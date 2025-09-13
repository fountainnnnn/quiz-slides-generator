
# Backend — Slides → Quiz Deck API

Windows-friendly build. Absolute download URLs returned from `/generate`.

## Setup

```bat
cd backend
python -m venv .venv
.\.venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
uvicorn src.app:app --host 0.0.0.0 --port 8000
```
