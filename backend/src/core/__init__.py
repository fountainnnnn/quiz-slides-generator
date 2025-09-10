# backend/src/core/__init__.py

# ---- Rendering ----
from .render import (
    pptx_to_png_windows_com,
    render_slides_to_images,
)

# ---- Notes / OCR→Markdown ----
from .notes import (
    extract_notes_and_alttext,
    bbox_to_xywh,
    normalize_lines,
    lines_to_markdown,
    ocr_slide_to_markdown,
    slides_to_clean_text,
)

# ---- OCR engines ----
from .ocr import (
    init_easyocr_temp,
    init_paddleocr,
    get_ocr_engine,
    ocr_image_easy,
    ocr_image_paddle,
)

# ---- Gemini QG helpers ----
from .gemini_qg import (
    configure_gemini,
    chunk_slides_for_qg,
    build_qg_prompt,
    safe_json_parse,
    generate_qa as generate_questions_with_gemini,  # back-compat alias
    generate_qa,
    explain_batch,
    infer_title,
)

# ---- PPTX export ----
from .pptx_export import build_qa_deck

# ---- Orchestration ----
from .pipeline import run_pipeline_end_to_end


__all__ = [
    # render
    "pptx_to_png_windows_com", "render_slides_to_images",
    # notes / ocr-md
    "extract_notes_and_alttext", "bbox_to_xywh", "normalize_lines",
    "lines_to_markdown", "ocr_slide_to_markdown", "slides_to_clean_text",
    # ocr
    "init_easyocr_temp", "init_paddleocr", "get_ocr_engine",
    "ocr_image_easy", "ocr_image_paddle",
    # gemini
    "configure_gemini", "chunk_slides_for_qg", "build_qg_prompt",
    "safe_json_parse", "generate_qa", "generate_questions_with_gemini",
    "explain_batch", "infer_title",
    # pptx
    "build_qa_deck",
    # pipeline
    "run_pipeline_end_to_end",
]
