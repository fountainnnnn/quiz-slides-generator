from .render import pptx_to_png_windows_com, render_slides_to_images
from .notes import extract_notes_and_alttext, bbox_to_xywh, normalize_lines, lines_to_markdown, ocr_slide_to_markdown, slides_to_clean_text
from .ocr import init_easyocr_temp, init_paddleocr, get_ocr_engine, ocr_image_easy, ocr_image_paddle
from .gemini_qg import configure_gemini, chunk_slides_for_qg, build_qg_prompt, safe_json_parse, generate_questions_with_gemini
from .pptx_export import _set_textbox_font, add_q_slide, add_a_slide, build_qa_deck

from .pipeline import run_pipeline_end_to_end
