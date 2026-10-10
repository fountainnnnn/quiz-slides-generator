# ---- Rendering ----
from .render import (
    pptx_to_png_windows_com,
    render_slides_to_images,
)

# ---- Notes / alt text ----
from .notes import extract_notes_and_alttext

# ---- Vision page reader (replaces EasyOCR) ----
from .vision import (
    VisionProvider,
    resolve_vision_provider,
    read_document,
    read_pages,
)

# ---- LLM QG helpers (DeepSeek) ----
from .llm_qg import (
    configure_deepseek,
    chunk_slides_for_qg,
    build_qg_prompt,
    safe_json_parse,
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
    # notes
    "extract_notes_and_alttext",
    # vision page reader
    "VisionProvider", "resolve_vision_provider", "read_document", "read_pages",
    # llm qg (DeepSeek)
    "configure_deepseek", "chunk_slides_for_qg", "build_qg_prompt",
    "safe_json_parse", "generate_qa",
    "explain_batch", "infer_title",
    # pptx
    "build_qa_deck",
    # pipeline
    "run_pipeline_end_to_end",
]
