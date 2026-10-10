# Quiz Slides Generator

Turns a lecture deck into a ready-to-present quiz deck. Upload a `.pptx` or `.pdf` and get back a PowerPoint with multiple-choice, theory, code-fill and fill-in-the-blank questions, each followed by its answer and a short explanation.

**Live demo:** https://ngyuhang.com/quiz-slide-generator/

## How it works

- **Reads every slide with a vision model:** slide text, code blocks exactly as written, and a one-line description of diagrams and charts. Without a vision key it falls back to the slides' embedded text.
- **Writes the questions with an LLM** (DeepSeek by default) in parallel batches, with strict JSON validation of every question.
- **Exports a PowerPoint** with question, answer and "Why?" slides.

## Keys

`DEEPSEEK_API_KEY` for question writing, plus one of `VISION_API_KEY`, `OPENROUTER_API_KEY` or `OPENAI_API_KEY` for slide reading (optional).

## Run locally

```bash
# API
cd backend
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export DEEPSEEK_API_KEY=...   # plus the optional keys listed above
CORS_ORIGINS=http://localhost:5173 uvicorn app:app --port 7860

# Frontend (second terminal)
cd frontend
python -m http.server 5173
```

Open http://localhost:5173. The page calls the API at `http://localhost:7860` by default; add `?api=<url>` to point it elsewhere.

## Tests

```bash
cd backend
pip install -r requirements-dev.txt 2>/dev/null || pip install pytest httpx
pytest
```

All model calls are mocked in the tests, so they run offline and cost nothing.

See [`backend/README.md`](backend/README.md) for every endpoint and environment variable.

## Known limitations

- Whole-slide images need LibreOffice; without it, slides are read from their text plus embedded pictures, so layout-only information can be missed.

## Project layout

- `backend/`: FastAPI service (the same code that runs the live demo)
- `frontend/`: static page (HTML, CSS, JavaScript)

## License

MIT
