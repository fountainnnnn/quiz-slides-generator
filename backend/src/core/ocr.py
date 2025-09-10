def init_easyocr_temp(lang_list=["en"], force_cpu=True):
    import easyocr
    tmp_dir = _mk_temp_dir("easyocr_")
    reader = easyocr.Reader(
        lang_list,
        gpu=not force_cpu,  # keep CPU by default
        model_storage_directory=str(tmp_dir),
        user_network_directory=str(tmp_dir),
    )
    reader._temp_model_dir = str(tmp_dir)  # mark for cleanup
    return reader

def init_paddleocr(lang="en"):
    # Only used if PaddleOCR is actually installed
    try:
        from paddleocr import PaddleOCR
    except Exception as e:
        raise RuntimeError(f"PaddleOCR not available: {e}")
    return PaddleOCR(use_angle_cls=True, lang=lang)

def get_ocr_engine(engine_name: str, lang: str):
    eng = (engine_name or "").lower()
    if eng == "paddle":
        try:
            ocr = init_paddleocr(lang)
            return ocr, True
        except Exception as e:
            print("[WARN] PaddleOCR unavailable; falling back to EasyOCR (CPU):", e)
    # default / fallback: EasyOCR on CPU
    return init_easyocr_temp([lang], force_cpu=True), False

def ocr_image_easy(reader, image):  # path | PIL | numpy
    if isinstance(image, (str, bytes)):
        res = reader.readtext(image, detail=1, paragraph=False)
    else:
        if isinstance(image, Image.Image):
            image = np.array(image.convert("RGB"))
        res = reader.readtext(image, detail=1, paragraph=False)
    return [{"bbox": b, "text": t, "conf": float(c)} for (b, t, c) in res]

def ocr_image_paddle(ocr, image):   # classic Paddle API
    res = ocr.ocr(image, cls=True)  # supports path or np array
    out = []
    for page in res:
        for bbox, (text, conf) in page:
            out.append({"bbox": bbox, "text": str(text).strip(), "conf": float(conf)})
    return out
