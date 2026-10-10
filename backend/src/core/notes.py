# -*- coding: utf-8 -*-
"""
Speaker notes + alt text extraction for PPTX slides.

Slide text itself is now read by the vision page reader (see vision.py),
which replaced the EasyOCR/OpenCV pipeline that used to live here.
"""

from typing import Any, Dict, List
from pathlib import Path

try:
    from pptx import Presentation  # python-pptx
except Exception:  # pragma: no cover
    Presentation = None  # type: ignore


# -----------------------------
# Notes + Alt text extraction
# -----------------------------
def extract_notes_and_alttext(pptx_path: str) -> Dict[int, Dict[str, Any]]:
    """
    Extract notes and alternative text per slide from a .pptx.

    Returns:
        {slide_index: {"notes": str, "alt_texts": List[str]}}
    """
    out: Dict[int, Dict[str, Any]] = {}
    p = Path(pptx_path)
    if p.suffix.lower() != ".pptx":
        return out
    if Presentation is None:
        return out

    prs = Presentation(str(p))
    for i, slide in enumerate(prs.slides, 1):
        # Notes
        notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text or ""
        except Exception:
            pass

        # Alt text (on shapes/pictures)
        alt_texts: List[str] = []
        for shape in slide.shapes:
            try:
                alt = getattr(shape, "alternative_text", "") or ""
                if alt and isinstance(alt, str):
                    alt = alt.strip()
                    if alt:
                        alt_texts.append(alt)
            except Exception:
                continue

        out[i] = {"notes": notes, "alt_texts": alt_texts}
    return out
