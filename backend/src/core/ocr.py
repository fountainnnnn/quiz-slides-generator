# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Any, Dict, List, Tuple, Optional
from pathlib import Path
import os
import tempfile

# --- Third-party (guarded) ---
try:
    import numpy as np  # type: ignore
except Exception:  # pragma: no cover
    np = None  # type: ignore

try:
    from PIL import Image  # type: ignore
except Exception:  # pragma: no cover
    Image = None  # type: ignore

# torch only for CUDA check (optional)
try:
    import torch  # type: ignore
except Exception:  # pragma: no cover
    torch = None  # type: ignore


# =========================
# Internal helpers / cache
# =========================
def _mk_temp_dir(prefix: str = "easyocr_") -> Path:
    """Create a temp directory we can optionally clean up later."""
    p = Path(tempfile.mkdtemp(prefix=prefix))
    return p

# Reuse EasyOCR readers so we don’t reload models for each call
# Keyed by (tuple(lang_list), use_gpu, storage_dir)
_EASYOCR_CACHE: Dict[Tuple[Tuple[str, ...], bool, str], Any] = {}


def _choose_storage_dir() -> Path:
    """
    Choose model cache directory.
    - Default: persistent under ~/.cache/easyocr_models  (faster restarts)
    - If SLIDES_OCR_EPHEMERAL=1: use a temp dir (clean each process)
    """
    if os.getenv("SLIDES_OCR_EPHEMERAL") == "1":
        return _mk_temp_dir("easyocr_")
    base = Path.home() / ".cache" / "easyocr_models"
    base.mkdir(parents=True, exist_ok=True)
    return base


def _gpu_allowed(force_cpu: bool) -> bool:
    if force_cpu:
        return False
    if torch is None:
        return False
    try:
        return bool(torch.cuda.is_available())
    except Exception:
        return False


# =========================
# Public API (unchanged)
# =========================
def init_easyocr_temp(lang_list: List[str] = ["en"], force_cpu: bool = True):
    """
    Initialize EasyOCR Reader with:
    - persistent model cache by default (override with SLIDES_OCR_EPHEMERAL=1)
    - GPU only if available AND force_cpu is False
    """
    import easyocr  # lazy import; raises if missing

    storage_dir = _choose_storage_dir()
    use_gpu = _gpu_allowed(force_cpu=force_cpu)

    key = (tuple(lang_list), use_gpu, str(storage_dir))
    if key in _EASYOCR_CACHE:
        return _EASYOCR_CACHE[key]

    reader = easyocr.Reader(
        lang_list,
        gpu=use_gpu,
        model_storage_directory=str(storage_dir),
        user_network_directory=str(storage_dir),
        download_enabled=True,   # ensure models can be fetched if missing
        verbose=False,
    )
    # mark a temp dir reference ONLY if ephemeral (so a caller can clean up)
    if os.getenv("SLIDES_OCR_EPHEMERAL") == "1":
        reader._temp_model_dir = str(storage_dir)  # type: ignore[attr-defined]
    _EASYOCR_CACHE[key] = reader
    return reader


def init_paddleocr(lang: str = "en"):
    """
    Initialize PaddleOCR (optional dependency). On Windows this often conflicts with protobuf.
    We keep it available behind a flag.
    """
    try:
        from paddleocr import PaddleOCR  # type: ignore
    except Exception as e:
        raise RuntimeError(f"PaddleOCR not available: {e}")
    return PaddleOCR(use_angle_cls=True, lang=lang)


def get_ocr_engine(engine_name: str, lang: str):
    """
    Returns (engine, use_paddle_bool)
    - "paddle" → PaddleOCR if available; else warn & fall back to EasyOCR (CPU)
    - default  → EasyOCR (CPU or GPU if available and not forced off)
    """
    eng = (engine_name or "").lower()
    if eng == "paddle":
        try:
            ocr = init_paddleocr(lang)
            return ocr, True
        except Exception as e:
            print("[WARN] PaddleOCR unavailable; falling back to EasyOCR (CPU):", e)
    # default / fallback: EasyOCR on CPU (unless CUDA available and not forced off)
    return init_easyocr_temp([lang], force_cpu=True), False


def ocr_image_easy(reader, image):  # str|bytes|PIL.Image|np.ndarray
    """
    Run EasyOCR on an image (path, bytes, PIL, or np array).
    Returns: List[{'bbox': [(x,y),...], 'text': str, 'conf': float}]
    """
    if reader is None:
        raise RuntimeError("EasyOCR reader is None.")
    if Image is None or np is None:
        raise RuntimeError("Pillow and numpy are required for EasyOCR.")

    try:
        # EasyOCR accepts file paths or numpy arrays
        if isinstance(image, (str, bytes, Path)):
            res = reader.readtext(str(image), detail=1, paragraph=False)
        else:
            if isinstance(image, Image.Image):
                image = np.array(image.convert("RGB"))
            # Slightly tuned params for slides (more structured text)
            res = reader.readtext(
                image,
                detail=1,
                paragraph=False,
                contrast_ths=0.05,
                adjust_contrast=0.5,
                text_threshold=0.6,
                low_text=0.3,
                width_ths=0.6,
                slope_ths=0.2,
                ycenter_ths=0.5,
                height_ths=0.6,
                mag_ratio=1.5,
            )
    except Exception as e:
        raise RuntimeError(f"EasyOCR failed: {e}")

    out: List[Dict[str, Any]] = []
    for item in res:
        # EasyOCR returns [(x,y), ...], text, conf
        try:
            bbox, text, conf = item
            text = (text or "").strip()
            if not text:
                continue
            out.append({"bbox": bbox, "text": text, "conf": float(conf)})
        except Exception:
            # unexpected tuple shape
            continue
    return out


def ocr_image_paddle(ocr, image):
    """
    Run PaddleOCR on an image (path or np array).
    Returns: List[{'bbox': [(x,y),...], 'text': str, 'conf': float}]
    """
    if ocr is None:
        raise RuntimeError("PaddleOCR engine is None.")

    try:
        res = ocr.ocr(image, cls=True)
    except Exception as e:
        raise RuntimeError(f"PaddleOCR failed: {e}")

    out: List[Dict[str, Any]] = []
    # Paddle returns [[ [ [x,y],...], (text, conf) ], ...] potentially wrapped per page
    try:
        for page in res:
            for det in page:
                bbox, meta = det
                text, conf = meta
                text = (text or "").strip()
                if not text:
                    continue
                out.append({"bbox": bbox, "text": str(text), "conf": float(conf)})
    except Exception:
        # Some Paddle builds return a single page (no nesting)
        for bbox, (text, conf) in res:
            text = (text or "").strip()
            if not text:
                continue
            out.append({"bbox": bbox, "text": str(text), "conf": float(conf)})
    return out
