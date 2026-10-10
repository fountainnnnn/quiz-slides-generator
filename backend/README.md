# Quiz Slides Generator API

FastAPI service behind the portfolio's `/quiz-slide-generator/` page. Upload a
lecture deck (`.pptx` or `.pdf`) and it returns a generated quiz deck (`.pptx`)
with MCQ, theory, code-fill and fill-in-the-blank questions plus explanations.

## How it works

1. **Page reading** (`src/core/vision.py`). Each page is rendered to PNG and a
   vision LLM transcribes it: slide text as Markdown, code blocks verbatim, and a
   one-line `[Diagram: ...]` description of diagrams and charts.
   - PDF pages are rendered with PyMuPDF.
   - PPTX slides are rendered whole when LibreOffice (`soffice`) or PowerPoint
     COM is available. The Docker image has neither, so in production each
     slide sends its text layer plus its embedded pictures (screenshots of code,
     diagrams) to the vision model. Speaker notes are appended verbatim.
   - Up to `VISION_CONCURRENCY` pages are read at once. Each call has a timeout
     and is retried with exponential backoff. Any page that still fails falls back
     to its embedded text layer, and so do pages past `VISION_MAX_PAGES`.
   - With **no vision key** the service still works in degraded mode: it uses
     the embedded text layer only (PyMuPDF `get_text` / python-pptx) and logs a
     warning. Scanned PDFs and text inside images are missed in this mode.
2. **Question generation** (`src/core/llm_qg.py`): DeepSeek through the OpenAI
   SDK, in parallel batches, with strict JSON validation.
3. **Export** (`src/core/pptx_export.py`): question, answer and "Why?" slides.

EasyOCR, OpenCV and torch were removed. The `ocr_engine` form field is still
accepted for compatibility but ignored.

## Endpoints

| Method | Path | Notes |
| --- | --- | --- |
| `POST` | `/generate` | multipart: `file` (pptx/pdf, max 15 MB), `total_questions`, `mix_mode` (`balanced`/`auto`/`custom`), `mcq_n`, `theory_n`, `codefill_n`, `fillblank_n`, `difficulty`, `include_explanations`, `include_thumbs`, `deck_title`, `dpi`, optional `openai_api_key` (DeepSeek key for this request only), `ocr_engine` (ignored) |
| `GET` | `/files/{filename}` | download a generated deck |
| `GET` | `/healthz` | health check used by Railway |

CORS allows only `https://ngyuhang.com`. `POST /generate` and `GET /files` are rate limited per client IP.

Uploads are saved under a per-request temp directory with a sanitised,
uuid-prefixed name and deleted after the request. A key sent in the form is
passed to the clients for that request and never written to the process
environment.

## Environment variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `DEEPSEEK_API_KEY` | (required) | text generation |
| `DEEPSEEK_BASE_URL` | `https://api.deepseek.com` | |
| `DEEPSEEK_MODEL` | `deepseek-v4-flash` | |
| `VISION_API_KEY` / `VISION_BASE_URL` / `VISION_MODEL` | | any OpenAI-compatible vision endpoint (checked first; model defaults to `gpt-4o-mini`) |
| `OPENROUTER_API_KEY` | | used if no `VISION_API_KEY`; base `https://openrouter.ai/api/v1`, model `qwen/qwen2.5-vl-72b-instruct` |
| `OPENAI_API_KEY` | | used if neither of the above; model `gpt-4o-mini` |
| `VISION_MODEL` | | overrides the vision model for any provider |
| `VISION_CONCURRENCY` | `4` | pages read in parallel |
| `VISION_PAGE_TIMEOUT_SECONDS` | `60` | per vision call |
| `VISION_MAX_RETRIES` | `2` | retries per page (backoff `VISION_RETRY_BACKOFF_SECONDS`, default `1.5`, doubled each time) |
| `VISION_MAX_PAGES` | `40` | pages sent to the vision model; later pages use the text layer |
| `VISION_MAX_IMAGES_PER_PAGE` | `4` | embedded pictures sent per PPTX slide |
| `QUIZ_GENERATOR_MAX_UPLOAD_BYTES` | `15728640` | upload cap |
| `QUIZ_GENERATOR_RATE_MAX` / `_RATE_WINDOW_SECONDS` | `4` / `60` | `/generate` rate limit |
| `QUIZ_GENERATOR_FILE_RATE_MAX` / `_FILE_RATE_WINDOW_SECONDS` | `60` / `60` | `/files` rate limit |
| `QUIZ_GENERATOR_BATCH_SIZE` / `QUIZ_GENERATOR_MAX_PARALLEL_CALLS` | `5` / `4` | question generation batching |
| `PUBLIC_QUIZ_GENERATOR_BASE_URL` | | base for returned download URLs |
| `PORT` | `7860` | |

## Run locally

```bash
python3.12 -m venv ../../.venv-quiz && source ../../.venv-quiz/bin/activate
pip install -r requirements.txt
export DEEPSEEK_API_KEY=...            # plus a vision key, e.g. OPENROUTER_API_KEY=...
uvicorn app:app --port 7860
curl localhost:7860/healthz
```

Or with Docker: `docker build -t quiz-generator . && docker run -p 7860:7860 -e DEEPSEEK_API_KEY=... quiz-generator`.

## Tests

```bash
pip install -r requirements.txt -r requirements-dev.txt
python -m pytest -q
```

The tests mock every LLM client (DeepSeek and vision) and clear provider keys,
so they make no network calls. They cover filename sanitising, provider
resolution order, text-layer fallback, the vision path, retries, page cap,
concurrency bound, `/healthz`, `/generate`, rate limiting and CORS.

`CORS_ORIGINS` (comma-separated, default `https://ngyuhang.com`) sets which browser origins may call the API. For local development with the static frontend, use `CORS_ORIGINS=http://localhost:5173`.
