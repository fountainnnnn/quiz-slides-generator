from typing import Any, Dict, List, Tuple, Optional
import os, io, json, re, tempfile
from pathlib import Path
from datetime import date

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from .vision import read_document


# =========================
# Extraction
# =========================
def _extract_lecture_text(path: Path, work_dir: Path, dpi: int, language: str) -> str:
    """
    Per-page Markdown from the vision page reader (falls back to the embedded
    text layer when no vision provider is configured), joined in page order.
    """
    pages = read_document(path, work_dir, dpi=dpi, language=language)
    return "\n\n".join(f"<!-- PAGE {i} -->\n{text}" for i, text in enumerate(pages, 1) if text.strip())


# =========================
# DeepSeek helpers
# =========================
def _get_api_key(explicit: Optional[str]) -> Optional[str]:
    return explicit or os.getenv("DEEPSEEK_API_KEY")


def _openai_generate_qa(
    text: str,
    total: int = 20,
    model_name: str = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-flash"),
    api_key: Optional[str] = None
) -> List[Tuple[str, str]]:
    """
    Simple Q/A generator (question, answer) using DeepSeek, with robust parsing.
    """
    key = _get_api_key(api_key)
    if not key:
        raise RuntimeError("DeepSeek API key missing. Set DEEPSEEK_API_KEY or pass an API key.")

    from openai import OpenAI
    client = OpenAI(
        api_key=key,
        base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
    )

    sys = (
        "You are a question generator. Read the provided lecture text and produce concise study questions "
        "with their short, correct answers. Return STRICT JSON: a list of objects with keys 'q' and 'a'. "
        "Avoid markdown or commentary."
    )
    prompt = f"{sys}\n\nTarget count: {total}\nLecture text:\n{text[:12000]}"

    resp = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        extra_body={"thinking": {"type": "disabled"}},
    )

    out: List[Tuple[str, str]] = []
    try:
        data = resp.choices[0].message.content or ""
        data = re.sub(r"^```(?:json)?\s*|\s*```$", "", data.strip(), flags=re.I | re.M)
        start = data.find("["); end = data.rfind("]")
        if start != -1 and end != -1:
            data = data[start:end+1]
        arr = json.loads(data)
        for item in arr:
            q = (item.get("q") or item.get("question") or "").strip()
            a = (item.get("a") or item.get("answer") or "").strip()
            if q and a:
                out.append((q, a))
    except Exception:
        # fallback parse
        text_resp = resp.choices[0].message.content or ""
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


def _openai_explain_batch(
    qa: List[Tuple[str, str]],
    text: str,
    model_name: str,
    api_key: Optional[str]
) -> List[str]:
    """
    2–3 sentence explanation per QA (same order). Robust JSON parsing.
    """
    key = _get_api_key(api_key)
    if not key:
        return [f"Explanation: {a[:200]}" for _, a in qa]

    from openai import OpenAI
    client = OpenAI(
        api_key=key,
        base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
    )

    pack = [{"q": q, "a": a} for q, a in qa]
    prompt = (
        "For each item, produce a short 2–3 sentence explanation (no formatting). "
        "Return STRICT JSON: a list of strings in the same order; its length must equal the input list.\n\n"
        f"LECTURE (truncated):\n{text[:8000]}\n\n"
        f"ITEMS:\n{json.dumps(pack, ensure_ascii=False)}"
    )

    resp = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        extra_body={"thinking": {"type": "disabled"}},
    )
    data = resp.choices[0].message.content or "[]"
    data = re.sub(r"^```(?:json)?\s*|\s*```$", "", data.strip(), flags=re.I | re.M)
    start, end = data.find("["), data.rfind("]")
    if start != -1 and end != -1:
        data = data[start:end+1]
    try:
        arr = json.loads(data)
        out = [str(x) for x in arr]
        if len(out) != len(qa):
            return [f"Explanation: {a[:200]}" for _, a in qa]
        return out
    except Exception:
        return [f"Explanation: {a[:200]}" for _, a in qa]


def _infer_title(
    lecture_text: str,
    filename_stem: str,
    model_name: str,
    api_key: Optional[str]
) -> str:
    """
    Try OpenAI for a concise 3–7 word title; otherwise heuristic or file name.
    """
    key = _get_api_key(api_key)
    if key:
        try:
            from openai import OpenAI
            client = OpenAI(
                api_key=key,
                base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
            )
            prompt = (
                "Give a concise 3–7 word title for this lecture text. "
                "Return ONLY the title, no quotes, no trailing punctuation.\n\n"
                f"{lecture_text[:8000]}"
            )
            resp = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": prompt}],
                extra_body={"thinking": {"type": "disabled"}},
            )
            t = (resp.choices[0].message.content or "").replace("\n", " ").strip(" .,:;\"'")
            if len(t.split()) >= 2:
                return t[:80]
        except Exception:
            pass
    # Heuristic fallback
    for line in lecture_text.splitlines():
        s = re.sub(r"[^A-Za-z0-9 :/\-\(\)\[\]]+", "", line).strip()
        if len(s.split()) >= 2 and len(s) >= 10:
            return s[:80]
    return filename_stem[:80] or "Auto Quiz"


# =========================
# Slide builder (improved)
# =========================
ACCENT_Q = RGBColor(37, 99, 235)    # blue
ACCENT_A = RGBColor(16, 185, 129)   # green
ACCENT_X = RGBColor(234, 179, 8)    # amber
FONT_NAME = "Calibri"

def _set_text_frame(tf, text: str, size_pt: int = 24, bold: bool = False,
                    color: Optional[RGBColor] = None, align=PP_ALIGN.LEFT):
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    for para in tf.paragraphs:
        para.space_after = Pt(6)
        para.space_before = Pt(0)
    tf.word_wrap = True


def _build_qa_deck(
    qa: List[Tuple[str, str]],
    title: str = "Auto Quiz",
    explanations: Optional[List[str]] = None,
    source_name: str = ""
) -> bytes:
    prs = Presentation()

    # Title slide
    s0 = prs.slides.add_slide(prs.slide_layouts[0])  # Title
    try:
        _set_text_frame(s0.shapes.title.text_frame, title, 44, True)
    except Exception:
        # if template differs, fallback to textbox
        tx = s0.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(1.5)).text_frame
        _set_text_frame(tx, title, 44, True)
    if len(s0.placeholders) > 1:
        sub = f"Generated quiz deck • {date.today().isoformat()}"
        if source_name:
            sub += f" • from {source_name}"
        try:
            _set_text_frame(s0.placeholders[1].text_frame, sub, 18)
        except Exception:
            pass

    total = len(qa)
    for i, (q, a) in enumerate(qa, start=1):
        # Question
        sq = prs.slides.add_slide(prs.slide_layouts[1])  # Title+Content
        try:
            _set_text_frame(sq.shapes.title.text_frame, f"Question {i}/{total}", 30, True, ACCENT_Q)
            _set_text_frame(sq.placeholders[1].text_frame, q, 28)
        except Exception:
            _set_textbox(sq, f"Question {i}/{total}", title=True, color=ACCENT_Q)
            _set_textbox(sq, q, top_inches=2.2)

        # Answer
        sa = prs.slides.add_slide(prs.slide_layouts[1])
        try:
            _set_text_frame(sa.shapes.title.text_frame, f"Answer {i}/{total}", 30, True, ACCENT_A)
            _set_text_frame(sa.placeholders[1].text_frame, a, 28)
        except Exception:
            _set_textbox(sa, f"Answer {i}/{total}", title=True, color=ACCENT_A)
            _set_textbox(sa, a, top_inches=2.2)

        # Explanation (optional)
        if explanations:
            sx = prs.slides.add_slide(prs.slide_layouts[1])
            text_x = explanations[i - 1]
            try:
                _set_text_frame(sx.shapes.title.text_frame, f"Why? {i}/{total}", 30, True, ACCENT_X)
                _set_text_frame(sx.placeholders[1].text_frame, text_x, 24)
            except Exception:
                _set_textbox(sx, f"Why? {i}/{total}", title=True, color=ACCENT_X)
                _set_textbox(sx, text_x, top_inches=2.2, size=24)

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()


def _set_textbox(slide, text: str, title: bool = False,
                 color: Optional[RGBColor] = None, top_inches: float = 1.5, size: int = 24):
    left, top, width, height = Inches(1), Inches(top_inches), Inches(8), Inches(4.5)
    tx = slide.shapes.add_textbox(left, top, width, height).text_frame
    tx.word_wrap = True
    p = tx.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(28 if title else size)
    if color:
        run.font.color.rgb = color
    if title:
        run.font.bold = True


# =========================
# Orchestration
# =========================
def run_pipeline_end_to_end(
    pptx_file,
    ocr_engine: str = "vision",   # deprecated: ignored, kept for API compatibility
    language: str = "en",
    prefer_com: bool = False,
    dpi: int = 180,
    openai_api_key: Optional[str] = None,
    model_name: str = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-flash"),
    total_questions: int = 20,
    mix_mode: str = "balanced",
    mcq_n: int = 0,
    theory_n: int = 0,
    codefill_n: int = 0,
    fillblank_n: int = 0,
    difficulty: str = "mixed",
    deck_title: str = "Auto Quiz",
    include_thumbs: bool = True,  # repurposed: include explanations when True
    source_name: Optional[str] = None,  # display name for the deck subtitle/title fallback
):
    """
    Returns: (pptx_bytes, zip_path, out_dir, msg)
    """
    # --- Save the upload with the correct extension in a request-scoped temp dir
    tmp_dir = Path(tempfile.mkdtemp(prefix="quizdeck_"))
    tmp_dir.mkdir(parents=True, exist_ok=True)

    ext = ""
    data: Optional[bytes] = None
    name_hint = getattr(pptx_file, "name", "") or ""

    if hasattr(pptx_file, "read"):  # file-like
        data = pptx_file.read()
        if name_hint:
            ext = Path(name_hint).suffix.lower()
    else:  # path-like
        p = Path(pptx_file)
        data = p.read_bytes()
        ext = p.suffix.lower()

    if not ext:
        if data and data.startswith(b"%PDF"):
            ext = ".pdf"
        elif data and data[:2] == b"PK":
            ext = ".pptx"

    if ext not in (".pptx", ".pdf"):
        raise ValueError(f"Unsupported file type: {ext or 'unknown'}. Upload .pptx or .pdf")

    tmp_in = tmp_dir / f"input_upload{ext}"
    tmp_in.write_bytes(data or b"")

    # --- Extract text (vision page reader, text-layer fallback)
    lecture_text = _extract_lecture_text(tmp_in, tmp_dir, dpi=dpi, language=language)

    if not lecture_text.strip():
        raise ValueError(
            "No text extracted from the document. If it is a scanned PDF, configure a vision "
            "provider (VISION_API_KEY, OPENROUTER_API_KEY or OPENAI_API_KEY)."
        )

    total_questions = max(1, min(int(total_questions), 50))
    mix_mode = (mix_mode or "balanced").lower()
    custom_counts = None
    if mix_mode == "custom":
        custom_counts = {
            "mcq": max(0, int(mcq_n)),
            "theory": max(0, int(theory_n)),
            "code_fill": max(0, int(codefill_n)),
            "fill_blank": max(0, int(fillblank_n)),
        }
        if sum(custom_counts.values()) != total_questions:
            raise ValueError("Custom question counts must add up to total_questions.")

    # --- Generate structured Q/A with question-type mix controls
    try:
        from .llm_qg import generate_qa as generate_structured_qa, infer_title as infer_structured_title
        from .pptx_export import build_qa_deck as build_structured_deck

        md_path = tmp_dir / "source.md"
        md_path.write_text(f"<!-- SLIDE 1 -->\n{lecture_text}", encoding="utf-8")
        qa_items = generate_structured_qa(
            per_slide_md_paths=[str(md_path)],
            total_questions=total_questions,
            mix=mix_mode,
            custom_counts=custom_counts,
            difficulty=difficulty,
            model_name=model_name,
            api_key=openai_api_key,
        )
        if not qa_items:
            raise ValueError("Q/A generation returned empty results.")

        filename_stem = Path(source_name or name_hint or tmp_in.name).stem
        auto_title = infer_structured_title(lecture_text, filename_stem, model_name, openai_api_key)
        final_title = deck_title if (deck_title and deck_title != "Auto Quiz") else auto_title
        out_path = tmp_dir / f"{filename_stem}_quizdeck.pptx"
        build_structured_deck(
            qa=qa_items,
            slide_images=[],
            out_path=str(out_path),
            deck_title=final_title,
            include_thumbnails=False,
            include_explanations=include_thumbs,
            source_name=filename_stem,
        )
        return out_path.read_bytes(), None, str(tmp_dir), "ok"
    except Exception as structured_error:
        # Keep a simpler Q/A fallback for older models or schema drift.
        print(f"[WARN] Structured quiz generation failed; falling back to simple Q/A: {structured_error}")

    qa = _openai_generate_qa(lecture_text, total=total_questions, model_name=model_name, api_key=openai_api_key)
    if not qa:
        raise ValueError("Q/A generation returned empty results.")

    # --- Title & Explanations (optional)
    filename_stem = Path(source_name or name_hint or tmp_in.name).stem
    auto_title = _infer_title(lecture_text, filename_stem, model_name, openai_api_key)
    final_title = deck_title if (deck_title and deck_title != "Auto Quiz") else auto_title

    explanations = _openai_explain_batch(qa, lecture_text, model_name, openai_api_key) if include_thumbs else None

    # --- Build deck (higher-quality)
    pptx_bytes = _build_qa_deck(qa, title=final_title, explanations=explanations, source_name=filename_stem)
    return pptx_bytes, None, str(tmp_dir), "ok"
