import os

import pytest
from fastapi.testclient import TestClient

from conftest import FakeChatClient, fake_quiz_llm, make_pdf

import app as app_module
from src.core import llm_qg, vision


@pytest.fixture
def client(monkeypatch, tmp_path):
    monkeypatch.setattr(app_module, "OUTPUT_DIR", tmp_path / "outputs")
    app_module.RATE_BUCKETS.clear()
    return TestClient(app_module.app)


@pytest.fixture
def fake_llms(monkeypatch):
    """Route every OpenAI-SDK client (DeepSeek and vision) to fakes and record the keys used."""
    seen = {"deepseek_keys": [], "vision": None}
    deepseek = FakeChatClient(fake_quiz_llm)

    def fake_openai(**kwargs):
        seen["deepseek_keys"].append(kwargs.get("api_key"))
        return deepseek

    vision_client = FakeChatClient(lambda kw: "# Hashing\n- Hash tables give O(1) average lookup\n[Diagram: buckets]")
    seen["vision"] = vision_client
    monkeypatch.setattr(llm_qg, "OpenAI", fake_openai)
    monkeypatch.setattr("openai.OpenAI", fake_openai)
    monkeypatch.setattr(vision, "make_vision_client", lambda provider, timeout: vision_client)
    return seen


# ---------- filename sanitising ----------
@pytest.mark.parametrize("raw, expected", [
    ("lecture.pdf", "lecture.pdf"),
    ("../../etc/passwd", "passwd"),
    ("..\\..\\windows\\evil.pptx", "evil.pptx"),
    ("/abs/path/My Slides (v2).pptx", "My_Slides_v2.pptx"),
    ("...hidden.pdf", "hidden.pdf"),
    ("", "upload"),
    (None, "upload"),
    ("..", "upload"),
    ("naïve résumé.pdf", "na_ve_r_sum.pdf"),
])
def test_safe_upload_name(raw, expected):
    assert app_module.safe_upload_name(raw) == expected


def test_unique_upload_name_has_uuid_prefix_and_no_separators():
    a = app_module.unique_upload_name("../x/../deck.pdf")
    b = app_module.unique_upload_name("../x/../deck.pdf")
    assert a != b
    assert a.endswith("_deck.pdf") and "/" not in a and "\\" not in a and ".." not in a


# ---------- endpoints ----------
def test_healthz(client):
    r = client.get("/healthz")
    assert r.status_code == 200 and r.json() == {"ok": True}


def test_generate_happy_path_with_mocked_llm_and_vision(client, fake_llms, tmp_path, monkeypatch):
    monkeypatch.setenv("DEEPSEEK_API_KEY", "server-key")
    monkeypatch.setenv("OPENAI_API_KEY", "vision-key")
    pdf = make_pdf(tmp_path / "in.pdf", ["Hash tables map keys to buckets for O(1) lookup."])

    with pdf.open("rb") as fh:
        r = client.post(
            "/generate",
            files={"file": ("../../evil/Lecture 1.pdf", fh, "application/pdf")},
            data={"total_questions": "1", "mix_mode": "balanced", "difficulty": "mixed",
                  "include_explanations": "true", "include_thumbs": "true", "ocr_engine": "easyocr"},
        )

    assert r.status_code == 200, r.text
    body = r.json()
    assert body["status"] == "ok"
    assert body["filename"].endswith("_Lecture_1_quizdeck.pptx")
    assert "/" not in body["filename"] and ".." not in body["filename"]
    assert body["url"].endswith(f"/files/{body['filename']}")
    assert fake_llms["vision"].calls, "vision reader should have been used"
    assert set(fake_llms["deepseek_keys"]) == {"server-key"}

    out_files = list((tmp_path / "outputs").iterdir())
    assert [f.name for f in out_files] == [body["filename"]]  # no upload copies left in the public dir

    dl = client.get(f"/files/{body['filename']}")
    assert dl.status_code == 200 and dl.content[:2] == b"PK"


def test_form_key_is_used_for_request_only(client, fake_llms, tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", ["Stacks are LIFO structures used for recursion."])
    with pdf.open("rb") as fh:
        r = client.post(
            "/generate",
            files={"file": ("in.pdf", fh, "application/pdf")},
            data={"total_questions": "1", "openai_api_key": "request-key"},
        )
    assert r.status_code == 200, r.text
    assert set(fake_llms["deepseek_keys"]) == {"request-key"}
    assert "DEEPSEEK_API_KEY" not in os.environ
    assert not fake_llms["vision"].calls  # no vision key configured -> text layer


def test_generate_rejects_wrong_type(client):
    r = client.post("/generate", files={"file": ("notes.txt", b"hi", "text/plain")})
    assert r.status_code == 400


def test_generate_missing_key(client, tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", ["text"])
    with pdf.open("rb") as fh:
        r = client.post("/generate", files={"file": ("in.pdf", fh, "application/pdf")})
    assert r.status_code == 400 and "DEEPSEEK_API_KEY" in r.json()["message"]


def test_rate_limit_still_enforced(client, monkeypatch):
    monkeypatch.setitem(app_module.RATE_LIMITS, ("POST", "/generate"), (2, 60))
    codes = [client.post("/generate", files={"file": ("x.txt", b"x", "text/plain")}).status_code for _ in range(3)]
    assert codes == [400, 400, 429]


def test_cors_only_allows_portfolio_origin(client):
    ok = client.options("/generate", headers={"Origin": "https://ngyuhang.com", "Access-Control-Request-Method": "POST"})
    bad = client.options("/generate", headers={"Origin": "https://evil.example", "Access-Control-Request-Method": "POST"})
    assert ok.headers.get("access-control-allow-origin") == "https://ngyuhang.com"
    assert "access-control-allow-origin" not in bad.headers
