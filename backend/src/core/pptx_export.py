from typing import Any, Dict, List, Tuple, Optional
def _set_textbox_font(text_frame, name="Calibri", size=24, bold=False, mono=False):
    for p in text_frame.paragraphs:
        for r in p.runs:
            r.font.name = "Consolas" if mono else name
            r.font.size = Pt(size)
            r.font.bold = bold

def add_q_slide(prs, title: str, body: str, mono: bool=False, thumbnail: Optional[str]=None):
    layout = prs.slide_layouts[1]  # Title & Content
    s = prs.slides.add_slide(layout)
    s.shapes.title.text = title
    tf = s.placeholders[1].text_frame
    tf.clear()
    for line in body.split("\n"):
        p = tf.add_paragraph() if tf.text else tf.paragraphs[0]
        p.text = line
        p.level = 0
    _set_textbox_font(tf, mono=mono)
    if thumbnail and Path(thumbnail).exists():
        left = Inches(9); top=Inches(1.5); height=Inches(3.5)
        try: s.shapes.add_picture(thumbnail, left, top, height=height)
        except Exception: pass
    return s

def add_a_slide(prs, title: str, body: str, mono: bool=False):
    layout = prs.slide_layouts[1]
    s = prs.slides.add_slide(layout)
    s.shapes.title.text = title
    tf = s.placeholders[1].text_frame
    tf.clear()
    for line in body.split("\n"):
        p = tf.add_paragraph() if tf.text else tf.paragraphs[0]
        p.text = line
        p.level = 0
    _set_textbox_font(tf, mono=mono)
    return s

def build_qa_deck(qa: List[Dict[str, Any]], slide_images: List[str], out_path: str, deck_title: str="Auto Q&A Deck", include_thumbnails: bool=False):
    prs = Presentation()
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = deck_title
    s0.placeholders[1].text = time.strftime("Generated on %Y-%m-%d %H:%M")

    for i, q in enumerate(qa, 1):
        typ = q["type"]
        src = int(q.get("source_slide_index", 1))
        thumb = slide_images[src-1] if include_thumbnails and 0 <= src-1 < len(slide_images) else None

        if typ == "mcq":
            opts = q["options"]
            opts_block = "\n".join([f"{chr(65+j)}. {opt}" for j,opt in enumerate(opts)])
            q_body = f"{q['question']}\n\n{opts_block}\n\n(Select one)"
            a_body = f"**Answer:** {q['answer']}\n\nExplanation: {q.get('explanation','')}"
            add_q_slide(prs, f"Q{i} (MCQ) — Slide {src}", q_body, mono=False, thumbnail=thumb)
            add_a_slide(prs, f"A{i}", a_body, mono=False)
        elif typ == "code_fill":
            q_body = f"{q['question']}\n\n{q.get('code','')}".strip()
            if q.get("options"): q_body += "\n\nOptions: " + ", ".join(q["options"])
            a_body = f"**Answer:** {q['answer']}\n\nExplanation: {q.get('explanation','')}"
            add_q_slide(prs, f"Q{i} (Code Fill) — Slide {src}", q_body, mono=True, thumbnail=thumb)
            add_a_slide(prs, f"A{i}", a_body, mono=True)
        else:
            q_body = q["question"]
            a_body = f"**Answer:** {q['answer']}\n\nExplanation: {q.get('explanation','')}"
            add_q_slide(prs, f"Q{i} (Theory) — Slide {src}", q_body, mono=False, thumbnail=thumb)
            add_a_slide(prs, f"A{i}", a_body, mono=False)

    prs.save(out_path)
    return out_path
