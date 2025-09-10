
import os
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from dotenv import load_dotenv

load_dotenv()

try:
    from src.core import run_pipeline_end_to_end
except Exception as e:
    raise RuntimeError(f"Failed to import pipeline from src.core: {e}")

app = FastAPI(title="Slides → Quiz Deck API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

@app.post("/generate")
async def generate(request: Request,
    file: UploadFile = File(...),
    ocr_engine: str = Form("easyocr"),
    language: str = Form("en"),
    prefer_com: bool = Form(False),
    dpi: int = Form(180),
    gemini_api_key: str = Form(""),
    model_name: str = Form("gemini-1.5-flash"),
    total_questions: int = Form(20),
    mix_mode: str = Form("balanced"),
    mcq_n: int = Form(0),
    theory_n: int = Form(0),
    codefill_n: int = Form(0),
    difficulty: str = Form("mixed"),
    deck_title: str = Form("Auto Quiz"),
    include_thumbs: bool = Form(True)
):
    content = await file.read()
    tmp_in = OUTPUT_DIR / f"upload_{file.filename}"
    tmp_in.write_bytes(content)

    key = gemini_api_key or os.getenv("GEMINI_API_KEY", "")
    if not key:
        return JSONResponse({"status":"error","message":"GEMINI_API_KEY missing. Provide via form or .env."}, status_code=400)
    os.environ["GEMINI_API_KEY"] = key

    try:
        result_pptx, zip_path, out_dir, msg = run_pipeline_end_to_end(
            pptx_file=tmp_in.open("rb"),
            ocr_engine=ocr_engine,
            language=language,
            prefer_com=prefer_com,
            dpi=dpi,
            gemini_api_key=key,
            model_name=model_name,
            total_questions=total_questions,
            mix_mode=mix_mode,
            mcq_n=mcq_n,
            theory_n=theory_n,
            codefill_n=codefill_n,
            difficulty=difficulty,
            deck_title=deck_title,
            include_thumbs=include_thumbs
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Pipeline error: {e}")

    if not result_pptx:
        raise HTTPException(status_code=500, detail=f"Generation failed: {msg}")

    # persist bytes and return absolute URL
    out_name = Path(file.filename).stem + "_quizdeck.pptx"
    out_path = OUTPUT_DIR / out_name
    try:
        data = Path(result_pptx).read_bytes()
    except Exception:
        data = result_pptx if isinstance(result_pptx, (bytes, bytearray)) else bytes(result_pptx)
    out_path.write_bytes(data)

    abs_url = str(request.url_for("get_file", filename=out_name))
    return {"status": "ok", "filename": out_name, "url": abs_url}

@app.get("/files/{filename}")
def get_file(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=filename)

@app.get("/healthz")
def healthz():
    return {"ok": True}
