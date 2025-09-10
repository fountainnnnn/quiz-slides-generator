
from typing import Any, Dict, List, Tuple, Optional
import os, io, json
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

import fitz  # PyMuPDF

# Optional EasyOCR (only used for bitmap PDFs/images)
try:
    import easyocr  # type: ignore
    _easyocr_reader = None
except Exception:
    easyocr = None
    _easyocr_reader = None

def _extract_text_from_pptx(pptx_path: Path) -> str:
    prs = Presentation(pptx_path)
    chunks = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                chunks.append(shape.text)
        if slide.has_notes_slide and slide.notes_slide and slide.notes_slide.notes_text_frame:
            nt = slide.notes_slide.notes_text_frame.text
            if nt:
                chunks.append(nt)
    return "\n\n".join(chunks)

def _extract_text_from_pdf(pdf_path: Path, use_ocr: bool = False, lang: str = "en") -> str:
    doc = fitz.open(pdf_path)
    parts, has_text = [], False
    for page in doc:
        txt = page.get_text().strip()
        if txt:
            has_text = True
            parts.append(txt)
    if has_text or not use_ocr:
        return "\n\n".join(parts)
    if easyocr is None:
        return "\n\n".join(parts)
    global _easyocr_reader
    if _easyocr_reader is None:
        _easyocr_reader = easyocr.Reader([lang], gpu=False)
    oc_parts = []
    for page in doc:
        pix = page.get_pixmap(dpi=200)
        import numpy as np, cv2
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        if img.shape[2] == 4:
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        results = _easyocr_reader.readtext(img, detail=0, paragraph=True)
        if results:
            oc_parts.append("\n".join(results))
    return "\n\n".join(oc_parts)

def _gemini_generate_qa(text: str, total: int = 20, model_name: str = "gemini-1.5-flash", api_key: Optional[str] = None) -> List[Tuple[str, str]]:
    import google.generativeai as genai
    if api_key:
        genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    sys = (
        "You are a question generator. Read the provided lecture text and produce concise study questions "
        "with their short, correct answers. Return STRICT JSON: a list of objects with keys 'q' and 'a'. "
        "Avoid markdown or commentary."
    )
    prompt = f"{sys}\n\nTarget count: {total}\nLecture text:\n{text[:12000]}"
    resp = model.generate_content(prompt)
    out: List[Tuple[str, str]] = []
    try:
        data = resp.text or ""
        start = data.find('['); end = data.rfind(']')
        if start != -1 and end != -1:
            data = data[start:end+1]
        arr = json.loads(data)
        for item in arr:
            q = (item.get("q") or item.get("question") or "").strip()
            a = (item.get("a") or item.get("answer") or "").strip()
            if q and a:
                out.append((q, a))
    except Exception:
        text_resp = getattr(resp, "text", "") or ""
        lines = [ln.strip() for ln in text_resp.splitlines() if ln.strip()]
        q, a = None, None
        for ln in lines:
            if ln.lower().startswith("q:"):
                if q and a:
                    out.append((q, a))
                q, a = ln[2:].strip(), ""
            elif ln.lower().startswith("a:") and q is not None:
                a = ln[2:].strip()
        if q and a:
            out.append((q, a))
    return out[:total]

def _add_textbox(slide, text: str, title=False):
    left, top, width, height = Inches(1), Inches(1.5), Inches(8), Inches(4.5)
    tx = slide.shapes.add_textbox(left, top, width, height).text_frame
    tx.word_wrap = True
    p = tx.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(28 if title else 24)

def _build_qa_deck(qa: List[Tuple[str, str]], title: str = "Auto Quiz") -> bytes:
    prs = Presentation()
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = title
    if len(s0.placeholders) > 1:
        s0.placeholders[1].text = "Generated quiz deck"
    for i, (q, a) in enumerate(qa, start=1):
        s1 = prs.slides.add_slide(prs.slide_layouts[5])
        _add_textbox(s1, f"Q{i}. {q}", title=True)
        s2 = prs.slides.add_slide(prs.slide_layouts[5])
        _add_textbox(s2, f"A{i}. {a}", title=True)
    bio = io.BytesIO(); prs.save(bio); return bio.getvalue()

def run_pipeline_end_to_end(
    pptx_file,
    ocr_engine: str = "easyocr",
    language: str = "en",
    prefer_com: bool = False,
    dpi: int = 180,
    gemini_api_key: Optional[str] = None,
    model_name: str = "gemini-1.5-flash",
    total_questions: int = 20,
    mix_mode: str = "balanced",
    mcq_n: int = 0,
    theory_n: int = 0,
    codefill_n: int = 0,
    difficulty: str = "mixed",
    deck_title: str = "Auto Quiz",
    include_thumbs: bool = True
):
    tmp_dir = Path(os.getenv("TMP", str(Path.cwd() / "tmp")))
    tmp_dir.mkdir(parents=True, exist_ok=True)

    # Get bytes + detect extension
    ext = ""
    data = None

    if hasattr(pptx_file, "read"):                 # file-like object (what app.py sends)
        data = pptx_file.read()
        name_hint = getattr(pptx_file, "name", "") or ""
        ext = Path(name_hint).suffix.lower()
    else:                                          # path-like was passed in
        p = Path(pptx_file)
        data = p.read_bytes()
        ext = p.suffix.lower()

    # If we still don't know, sniff by magic bytes
    if not ext:
        if data.startswith(b"%PDF"):
            ext = ".pdf"
        elif data[:2] == b"PK":  # .pptx is a zip
            ext = ".pptx"

    if ext not in (".pptx", ".pdf"):
        raise ValueError(f"Unsupported file type: {ext or 'unknown'}. Upload .pptx or .pdf")

    # Write temp file with the correct extension
    tmp_in = tmp_dir / f"input_upload{ext}"
    tmp_in.write_bytes(data)

    if ext == ".pptx":
        lecture_text = _extract_text_from_pptx(tmp_in)
    elif ext == ".pdf":
        lecture_text = _extract_text_from_pdf(tmp_in, use_ocr=(ocr_engine=="easyocr"), lang=language)
    else:
        raise ValueError("Unsupported file type. Upload .pptx or .pdf")
    if not lecture_text.strip():
        raise ValueError("No text extracted from the document. Check OCR/inputs.")
    qa = _gemini_generate_qa(lecture_text, total=total_questions, model_name=model_name, api_key=gemini_api_key)
    if not qa:
        raise ValueError("Q/A generation returned empty results.")
    pptx_bytes = _build_qa_deck(qa, title=deck_title)
    return pptx_bytes, None, None, "ok"
