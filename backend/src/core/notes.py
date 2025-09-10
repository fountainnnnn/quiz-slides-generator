from typing import Any, Dict, List, Tuple, Optional
def extract_notes_and_alttext(pptx_path: str) -> Dict[int, Dict[str, Any]]:
    data = {}
    if Path(pptx_path).suffix.lower() != ".pptx":
        return data
    prs = Presentation(pptx_path)
    for i, slide in enumerate(prs.slides, 1):
        notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text or ""
        except Exception:
            pass
        alt_texts = []
        for shape in slide.shapes:
            try:
                alt = getattr(shape, "alternative_text", "") or ""
                if alt.strip():
                    alt_texts.append(alt.strip())
            except Exception:
                pass
        data[i] = {"notes": notes, "alt_texts": alt_texts}
    return data
def bbox_to_xywh(b):
    xs = [p[0] for p in b]; ys = [p[1] for p in b]
    x, y = min(xs), min(ys)
    w, h = max(xs) - x, max(ys) - y
    return x, y, w, h

def normalize_lines(items, y_merge=10):
    enriched = []
    for it in items:
        xywh = bbox_to_xywh(it["bbox"])
        enriched.append({**it, "xywh": xywh})
    enriched = sorted(enriched, key=lambda d: (round(d["xywh"][1]/y_merge), d["xywh"][0]))

    merged, current, last_row = [], None, None
    for it in enriched:
        x,y,w,h = it["xywh"]
        text = (it.get("text") or "").strip()
        if not text: continue
        row = round(y / y_merge)
        if current is None:
            current = {"row": row, "x": x, "texts": [text], "xs": [x], "ys": [y], "conf": [it.get("conf",1.0)]}
            last_row = row
        else:
            if row == last_row:
                current["texts"].append(text)
                current["xs"].append(x); current["ys"].append(y); current["conf"].append(it.get("conf",1.0))
            else:
                merged.append(current)
                current = {"row": row, "x": x, "texts": [text], "xs": [x], "ys": [y], "conf": [it.get("conf",1.0)]}
                last_row = row
    if current: merged.append(current)

    lines = []
    for m in merged:
        line_text = re.sub(r"\s+", " ", " ".join(m["texts"]).strip())
        if not line_text: continue
        avg_conf = float(np.mean(m["conf"])) if m["conf"] else 1.0
        lines.append({"text": line_text, "left": min(m["xs"]), "y": float(np.median(m["ys"])), "conf": avg_conf})
    return lines

def lines_to_markdown(lines, conf_threshold=0.45):
    lines = [l for l in lines if l["conf"] >= conf_threshold and l["text"]]
    if not lines: return ""
    lefts = np.array([l["left"] for l in lines], dtype=float)
    base = float(np.percentile(lefts, 10))
    indent_unit = max(10.0, float(np.percentile(lefts - base, 75) / 2.0))
    md = []
    first = lines[0]["text"]
    title = None
    if len(first) <= 80 and first[:1].upper()==first[:1] and not first.endswith("."):
        title = "# " + first; lines = lines[1:]
    if title: md.append(title)
    for l in lines:
        depth = int(max(0, round((l["left"] - base) / indent_unit)))
        bullet = "  " * depth + "- " + l["text"]
        md.append(bullet)
    return "\n".join(md)
def ocr_slide_to_markdown(image_path: str, ocr_engine, use_paddle: bool = False, max_width: int = 1600) -> str:
    img = Image.open(image_path).convert("RGB")
    if img.width > max_width:
        ratio = max_width / img.width
        img = img.resize((max_width, int(img.height * ratio)))
    img_np = np.array(img)
    items = ocr_image_paddle(ocr_engine, img_np) if use_paddle else ocr_image_easy(ocr_engine, img_np)
    lines = normalize_lines(items)
    return lines_to_markdown(lines)

def slides_to_clean_text(
    pptx_path: str,
    out_dir: str = "slide_text_output",
    ocr_engine_name: str = "easyocr",
    lang: str = "en",
    prefer_windows_com: bool = True,
    dpi: int = 180
) -> Dict[str, Any]:
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    slides_dir = out / "slides"; slides_dir.mkdir(exist_ok=True)

    ext = Path(pptx_path).suffix.lower()

    # Render images
    if ext == ".pdf":
        import fitz  # PyMuPDF
        doc = fitz.open(pptx_path)
        slide_imgs = []
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            img_path = slides_dir / f"slide_{i:03d}.png"
            pix.save(str(img_path))
            slide_imgs.append(img_path)
    else:
        slide_imgs = render_slides_to_images(
            pptx_path, out_dir=str(slides_dir), dpi=dpi, prefer_windows_com=prefer_windows_com
        )

    meta = extract_notes_and_alttext(pptx_path) if ext == ".pptx" else {}

    ocr_engine, use_paddle = get_ocr_engine(ocr_engine_name, lang)

    per_slide_md_paths = []; all_md_blocks = []
    try:
        for idx, img_path in enumerate(tqdm(sorted(slide_imgs), desc="OCR slides"), 1):
            md = ocr_slide_to_markdown(str(img_path), ocr_engine, use_paddle=use_paddle)
            notes = meta.get(idx, {}).get("notes", "").strip()
            alt_texts = meta.get(idx, {}).get("alt_texts", [])
            if notes: md += f"\n\n> Notes:\n> {notes}"
            if alt_texts: md += "\n\n> Alt text:\n" + "\n".join([f"> - {t}" for t in alt_texts])
            if not md.strip(): md = f"# Slide {idx}\n- (No text detected)"
            slide_md_path = out / f"slide_{idx:03d}.md"; slide_md_path.write_text(md, encoding="utf-8")
            per_slide_md_paths.append(slide_md_path)
            all_md_blocks.append(f"<!-- SLIDE {idx} -->\n{md}\n")
    finally:
        # Clean up temp OCR cache if we created one
        cleanup_ocr_cache(ocr_engine)

    concat_path = out / "slides_concatenated.txt"
    concat_path.write_text("\n\n".join(all_md_blocks), encoding="utf-8")

    # Zip
    zip_path = out / "slide_text_outputs.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(concat_path, arcname=concat_path.name)
        for mdp in per_slide_md_paths: zf.write(mdp, arcname=mdp.name)
        for img in slide_imgs: zf.write(img, arcname=f"slides/{Path(img).name}")

    return {
        "images": [str(p) for p in slide_imgs],
        "per_slide_markdown": [str(p) for p in per_slide_md_paths],
        "concat_txt": str(concat_path),
        "zip": str(zip_path),
        "out_dir": str(out)
    }
