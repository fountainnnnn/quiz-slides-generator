import json
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

SERVICE_ROOT = Path(__file__).resolve().parents[1]
if str(SERVICE_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVICE_ROOT))

PROVIDER_ENV = (
    "VISION_API_KEY", "VISION_BASE_URL", "VISION_MODEL",
    "OPENROUTER_API_KEY", "OPENAI_API_KEY", "DEEPSEEK_API_KEY",
)


@pytest.fixture(autouse=True)
def clean_env(monkeypatch):
    """No test may see real keys or make a live call."""
    for name in PROVIDER_ENV:
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("VISION_RETRY_BACKOFF_SECONDS", "0")
    yield


def make_pdf(path: Path, pages):
    import fitz

    doc = fitz.open()
    for text in pages:
        page = doc.new_page()
        if text:
            page.insert_text((72, 72), text)
    doc.save(str(path))
    doc.close()
    return path


def _response(content: str):
    return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))])


class FakeChatClient:
    """Mimics openai.OpenAI().chat.completions.create; records calls, never touches the network."""

    def __init__(self, handler):
        self.calls = []
        self._handler = handler
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs):
        self.calls.append(kwargs)
        return _response(self._handler(kwargs))


def fake_quiz_llm(kwargs):
    """Answers QG / title / explanation prompts with valid canned output."""
    prompt = kwargs["messages"][-1]["content"]
    if isinstance(prompt, str) and "Return JSON only" in prompt:
        return json.dumps({"questions": [{
            "type": "mcq",
            "question": "Which structure gives O(1) average lookup?",
            "options": ["Hash table", "Linked list", "Array scan", "Stack"],
            "answer": "Hash table",
            "explanation": "Hashing maps keys directly to buckets.",
            "source_slide_index": 1,
        }]})
    if isinstance(prompt, str) and "concise 3" in prompt:
        return "Data Structures Basics"
    return "[]"
