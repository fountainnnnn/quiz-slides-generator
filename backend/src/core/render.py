# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List
import sys
import shutil
import subprocess

# ---------------------------
# Utilities
# ---------------------------
def _pdf_to_png_pymupdf(pdf_path: str | Path, out_dir: str | Path, dpi: int) -> List[Path]:
    """
    Rasterize a PDF to PNG pages using PyMuPDF at a target DPI.
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    import fitz  # PyMuPDF
    doc = fitz.open(str(pdf_path))
    imgs: List[Path] = []
    # DPI -> zoom factor (72pt = 1")
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=mat)
        p = out / f"slide_{i:03d}.png"
        pix.save(str(p))
        imgs.append(p)
    if not imgs:
        raise RuntimeError("PDF rasterization produced no images.")
    return imgs


# ---------------------------
# Windows COM (PowerPoint) → PDF → PNG
# ---------------------------
def _pptx_to_pdf_windows_com(pptx_path: str | Path, pdf_path: str | Path) -> None:
    """
    Export PPTX to PDF using PowerPoint COM automation (Windows only).
    """
    from win32com.client import Dispatch  # type: ignore

    pptx_path = str(Path(pptx_path).resolve())
    pdf_path = str(Path(pdf_path).resolve())

    powerpoint = Dispatch("PowerPoint.Application")
    powerpoint.Visible = 0  # hidden
    try:
        pres = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        # 2 = ppFixedFormatTypePDF, 2 = ppFixedFormatIntentPrint
        pres.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            FixedFormatType=2,
            Intent=2,
            FrameSlides=False,
            RangeType=1,  # ppPrintAll
        )
        pres.Close()
    finally:
        powerpoint.Quit()


def pptx_to_png_windows_com(pptx_path: str, out_dir: str) -> List[Path]:
    """
    Backward-compatible wrapper name.
    Now exports PPTX → PDF with COM, then rasterizes to PNG at a good DPI (220).
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    tmp_pdf = out / "_export_tmp.pdf"
    _pptx_to_pdf_windows_com(pptx_path, tmp_pdf)
    try:
        return _pdf_to_png_pymupdf(tmp_pdf, out, dpi=220)
    finally:
        try:
            tmp_pdf.unlink(missing_ok=True)  # type: ignore[arg-type]
        except Exception:
            pass


# ---------------------------
# LibreOffice headless fallback
# ---------------------------
def _pptx_to_pdf_libreoffice(pptx_path: str | Path, out_pdf_path: str | Path) -> None:
    """
    Convert PPTX → PDF using LibreOffice in headless mode, if installed.
    Works on Windows/macOS/Linux wherever 'soffice' is available.
    """
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "LibreOffice ('soffice') not found. Install LibreOffice or provide a PDF instead."
        )

    out_pdf_path = Path(out_pdf_path)
    tmp_dir = out_pdf_path.parent
    tmp_dir.mkdir(parents=True, exist_ok=True)

    # Convert to PDF in the target directory
    cmd = [
        soffice, "--headless", "--norestore", "--invisible", "--nodefault", "--nolockcheck",
        "--convert-to", "pdf",
        "--outdir", str(tmp_dir),
        str(Path(pptx_path).resolve())
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {proc.stderr or proc.stdout}")

    # LibreOffice writes <basename>.pdf into outdir
    produced = Path(tmp_dir) / (Path(pptx_path).stem + ".pdf")
    if not produced.exists():
        raise RuntimeError("LibreOffice did not produce a PDF as expected.")
    # Move/rename to desired path
    if out_pdf_path != produced:
        produced.replace(out_pdf_path)


# ---------------------------
# Public entry point
# ---------------------------
def render_slides_to_images(
    pptx_path: str,
    out_dir: str,
    dpi: int = 180,
    prefer_windows_com: bool = True
) -> List[Path]:
    """
    Render PPTX/PDF to per-page PNG images.
    Strategy:
      - If input is PDF → rasterize directly (PyMuPDF).
      - If PPTX:
          * On Windows + prefer_windows_com: PowerPoint COM → PDF → PNG (DPI respected).
          * Else if LibreOffice available: soffice → PDF → PNG.
          * Else: instruct user to upload PDF.
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    ext = Path(pptx_path).suffix.lower()

    # Direct PDF path
    if ext == ".pdf":
        return _pdf_to_png_pymupdf(pptx_path, out, dpi=dpi)

    # PPTX path
    if sys.platform.startswith("win") and prefer_windows_com:
        try:
            # Export via COM to PDF, then rasterize with desired DPI
            tmp_pdf = out / "_export_tmp.pdf"
            _pptx_to_pdf_windows_com(pptx_path, tmp_pdf)
            try:
                return _pdf_to_png_pymupdf(tmp_pdf, out, dpi=dpi)
            finally:
                try:
                    tmp_pdf.unlink(missing_ok=True)  # type: ignore[arg-type]
                except Exception:
                    pass
        except Exception as e:
            print("[WARN] PowerPoint COM export failed; will try LibreOffice if present:", e)

    # LibreOffice headless fallback (cross-platform)
    try:
        tmp_pdf = out / "_export_tmp.pdf"
        _pptx_to_pdf_libreoffice(pptx_path, tmp_pdf)
        try:
            return _pdf_to_png_pymupdf(tmp_pdf, out, dpi=dpi)
        finally:
            try:
                tmp_pdf.unlink(missing_ok=True)  # type: ignore[arg-type]
            except Exception:
                pass
    except Exception as e:
        raise RuntimeError(
            "Could not render PPTX on this environment. "
            "Options:\n"
            "  1) Install LibreOffice and ensure 'soffice' is on PATH,\n"
            "  2) Run on Windows with PowerPoint installed (prefer_windows_com=True), or\n"
            "  3) Export your deck to PDF on the client and upload that.\n"
            f"Details: {e}"
        )
