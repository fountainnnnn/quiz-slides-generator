def pptx_to_png_windows_com(pptx_path: str, out_dir: str):
    import win32com.client as win32
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    powerpoint = win32.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 0
    pres = powerpoint.Presentations.Open(str(Path(pptx_path).resolve()))
    try:
        pres.Export(str(Path(out_dir).resolve()), "PNG")
    finally:
        pres.Close(); powerpoint.Quit()
    normalized = []
    for i, p in enumerate(sorted(Path(out_dir).glob("Slide*.PNG")), 1):
        newp = Path(out_dir) / f"slide_{i:03d}.png"
        p.rename(newp); normalized.append(newp)
    if not normalized:
        raise RuntimeError("PowerPoint Export produced no images.")
    return normalized

def render_slides_to_images(pptx_path: str, out_dir: str, dpi: int = 180, prefer_windows_com: bool = True):
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    ext = Path(pptx_path).suffix.lower()
    if ext == ".pdf":
        import fitz  # PyMuPDF
        doc = fitz.open(str(pptx_path))
        imgs = []
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            p = out / f"slide_{i:03d}.png"
            pix.save(str(p)); imgs.append(p)
        return imgs
    # PPTX
    if sys.platform.startswith("win") and prefer_windows_com:
        try:
            return sorted(pptx_to_png_windows_com(pptx_path, out_dir))
        except Exception as e:
            print("[WARN] PowerPoint COM export failed:", e)
    # If we get here and still have a PPTX, we don't have a universal no-COM fallback.
    raise RuntimeError("Could not render PPTX. Ensure MS PowerPoint is installed/enabled. Alternatively, export your deck to PDF and re-run.")
