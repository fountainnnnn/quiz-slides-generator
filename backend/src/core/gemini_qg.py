# backend/src/core/gemini_qg.py

from typing import Any, Dict, List, Tuple, Optional
from pathlib import Path
import os, math, re, json as _json

GEMINI_DEFAULT_MODEL = "gemini-1.5-flash"

# ------------------------------------------------------------
# Gemini configuration (env only — no embedded default key)
# ------------------------------------------------------------
def configure_gemini(api_key: Optional[str] = None):
    """
    Configure google-generativeai with a required API key.

    Precedence:
      1) function arg `api_key`
      2) env var GEMINI_API_KEY
      3) env var GOOGLE_API_KEY

    If none is set, raises a clear error.
    """
    key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key:
        raise RuntimeError(
            "Gemini API key not found. Create a key in Google AI Studio and set it on the backend "
            "via environment variable GEMINI_API_KEY (recommended) or pass gemini_api_key in the request."
        )
    import google.generativeai as genai
    genai.configure(api_key=key)
    return genai


# ------------------------------------------------------------
# Prompting
# ------------------------------------------------------------
QG_SYSTEM = (
    "You are an expert quiz generator for advanced university students. "
    "You will receive slide content as tidy markdown blocks, each tagged with its slide number. "
    "Generate rigorous, challenging questions grounded ONLY in the provided content.\n\n"
    "Question types:\n"
    "- \"mcq\": 4–5 options; avoid obvious distractors; make options plausible but only one correct. "
    "Include a short 'explanation' with reasoning why the correct option is right and others are wrong.\n"
    "- \"theory\": short-answer/conceptual; require precise definitions, derivations, or reasoning; "
    "put the ideal answer in 'answer' and a detailed 'explanation'.\n"
    "- \"code_fill\": provide a prompt + a code block with blanks like `___`; "
    "blanks should test understanding of syntax, logic, or algorithm steps. "
    "Put the exact filled line(s) in 'answer'; include a reasoning 'explanation'.\n"
    "- \"fill_blank\": prose/sentence(s) with blanks `___`; "
    "blanks should be non-trivial concepts or technical terms. "
    "Put the exact fill(s) in 'answer' (string or list); include an 'explanation'.\n\n"
    "Rules:\n"
    "- Always include 'source_slide_index' (1-based index matching the slide tag).\n"
    "- Make questions challenging: test reasoning, synthesis, and nuance, not just recall. "
    "Combine ideas where possible.\n"
    "- Only use facts derivable from the provided slides; do not fabricate.\n"
    "- Difficulty guidance is provided; align output but avoid triviality.\n\n"
    "STRICT OUTPUT FORMAT: Return ONLY JSON — a list of objects, no prose, no markdown fences. "
    "Allowed keys per item: "
    "['type','question','options','answer','explanation','source_slide_index','code','text_with_blanks']"
)


def chunk_slides_for_qg(slide_md_paths: List[str], max_chars_per_chunk: int = 8000):
    chunks: List[Tuple[List[int], str]] = []
    buf: List[str] = []
    buf_len = 0
    idxs: List[int] = []
    for i, p in enumerate(slide_md_paths, 1):
        t = Path(p).read_text(encoding="utf-8")
        block = f"\n\n<!-- SLIDE {i} -->\n{t}\n"
        if buf_len + len(block) > max_chars_per_chunk and buf:
            chunks.append((idxs, "".join(buf)))
            buf, buf_len, idxs = [], 0, []
        buf.append(block); buf_len += len(block); idxs.append(i)
    if buf:
        chunks.append((idxs, "".join(buf)))
    return chunks

def build_qg_prompt(slide_block: str, want_counts: Dict[str,int], difficulty: str):
    total = max(1, sum(max(0, v) for v in want_counts.values()))
    mix_desc = ", ".join([f"{k}:{v}" for k,v in want_counts.items() if v>0]) or "auto"
    return f"""{QG_SYSTEM}

Difficulty target: {difficulty}.
Generate exactly {total} items with mix {mix_desc}. If the content cannot support a type, reallocate to others.

Slide content:
{slide_block}

JSON only."""
# ------------------------------------------------------------


# ------------------------------------------------------------
# Parsing & validation
# ------------------------------------------------------------
def safe_json_parse(s: str) -> List[Dict[str, Any]]:
    """Parse JSON robustly (handles ```json fences, leading/trailing noise)."""
    if not s:
        return []
    # strip fences
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s.strip(), flags=re.I|re.M)
    # try full parse
    try:
        obj = _json.loads(s)
        return obj if isinstance(obj, list) else []
    except Exception:
        pass
    # find first [...] block
    m = re.search(r"\[[\s\S]*\]", s)
    if m:
        try:
            obj = _json.loads(m.group(0))
            return obj if isinstance(obj, list) else []
        except Exception:
            return []
    return []

def _coerce_answer_to_str(ans: Any) -> str:
    if ans is None:
        return ""
    if isinstance(ans, list):
        return ", ".join([str(a).strip() for a in ans if str(a).strip()])
    return str(ans).strip()

def _clean_and_validate(items: List[Dict[str, Any]], idxs_fallback: List[int]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    seen: set = set()
    for raw in items:
        if not isinstance(raw, dict):
            continue
        t = str(raw.get("type", "")).lower().strip()
        if t not in {"mcq", "theory", "code_fill", "fill_blank"}:
            continue

        # normalize fields
        question = (raw.get("question") or raw.get("text_with_blanks") or "").strip()
        answer = _coerce_answer_to_str(raw.get("answer"))
        if not question or not answer:
            continue

        # options / code normalization
        opts: List[str] = []
        code = ""
        if t == "mcq":
            opts_raw = raw.get("options", [])
            if not isinstance(opts_raw, list):
                continue
            opts = [str(o).strip() for o in opts_raw if str(o).strip()]
            if len(opts) < 3:
                continue
            # allow A/B/C/D mapping
            if answer not in opts and answer.upper() in ["A","B","C","D","E"]:
                idx = ord(answer.upper()) - 65
                if 0 <= idx < len(opts):
                    answer = opts[idx]
            if answer not in opts:
                continue
        elif t == "code_fill":
            code = (raw.get("code") or "").strip()
            opts = [str(o).strip() for o in (raw.get("options") or []) if str(o).strip()]
        else:
            # theory / fill_blank: no options by default
            opts = []

        exp = (raw.get("explanation") or "").strip()
        try:
            src = int(raw.get("source_slide_index", 0)) or (idxs_fallback[0] if idxs_fallback else 1)
        except Exception:
            src = idxs_fallback[0] if idxs_fallback else 1

        # dedupe by (type, question[:160], answer[:160])
        k = (t, question[:160], answer[:160])
        if k in seen:
            continue
        seen.add(k)

        item: Dict[str, Any] = {
            "type": t,
            "question": question[:1200],
            "options": opts[:6],
            "answer": answer[:1200],
            "explanation": exp[:1200],
            "source_slide_index": src,
        }
        if code:
            item["code"] = code[:4000]
        out.append(item)
    return out
# ------------------------------------------------------------


# ------------------------------------------------------------
# Public helpers
# ------------------------------------------------------------
def generate_qa(
    per_slide_md_paths: List[str],
    total_questions: int = 20,
    mix: str = "auto",  # "auto" | "balanced" | "custom"
    custom_counts: Optional[Dict[str,int]] = None,
    difficulty: str = "mixed",
    model_name: str = GEMINI_DEFAULT_MODEL,
    api_key: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Generate questions using Gemini (key from env or arg). Strict JSON parsing with validation.
    Supports types: mcq, theory, code_fill, fill_blank.
    """
    genai = configure_gemini(api_key)
    model = genai.GenerativeModel(model_name)

    # desired mix
    if mix == "custom" and custom_counts:
        desired = {
            "mcq":      max(0, int(custom_counts.get("mcq", 0))),
            "theory":   max(0, int(custom_counts.get("theory", 0))),
            "code_fill":max(0, int(custom_counts.get("code_fill", 0))),
            "fill_blank":max(0, int(custom_counts.get("fill_blank", 0))),
        }
        if sum(desired.values()) <= 0:
            desired = {"mcq": total_questions}
    elif mix == "balanced":
        q4 = max(1, total_questions // 4)
        desired = {
            "mcq": q4,
            "theory": q4,
            "code_fill": q4,
            "fill_blank": total_questions - 3*q4
        }
    else:
        # auto: bias toward MCQ but include others
        mcq = math.ceil(total_questions * 0.45)
        theory = max(0, math.ceil(total_questions * 0.25))
        code_fill = max(0, math.ceil(total_questions * 0.15))
        fill_blank = max(0, total_questions - mcq - theory - code_fill)
        desired = {"mcq": mcq, "theory": theory, "code_fill": code_fill, "fill_blank": fill_blank}

    chunks = chunk_slides_for_qg(per_slide_md_paths, max_chars_per_chunk=8000)
    remaining = total_questions
    results: List[Dict[str, Any]] = []

    for idxs, block in chunks:
        if remaining <= 0:
            break

        # proportional allocation by number of slides in chunk
        weight = max(1, len(idxs))
        per_chunk = max(
            1,
            min(remaining, math.ceil(total_questions * (weight / max(1, len(per_slide_md_paths)))))
        )

        # scale desired mix into this chunk
        want_counts = desired.copy()
        s = sum(want_counts.values()) or 1
        for k in want_counts:
            want_counts[k] = max(0, round(want_counts[k] * per_chunk / s))
        # fix rounding drift
        drift = per_chunk - sum(want_counts.values())
        for k in ["mcq", "theory", "code_fill", "fill_blank"]:
            if drift == 0:
                break
            want_counts[k] += 1
            drift -= 1

        prompt = build_qg_prompt(block, want_counts, difficulty)
        resp = model.generate_content(prompt)
        text = getattr(resp, "text", None) or (
            resp.candidates[0].content.parts[0].text if getattr(resp, "candidates", None) else ""
        )

        raw = safe_json_parse(text)
        clean = _clean_and_validate(raw, idxs)

        # cap per chunk, append
        clean = clean[:per_chunk]
        results.extend(clean)
        remaining -= len(clean)

    return results[:total_questions]


def explain_batch(
    qa: List[Tuple[str, str]] | List[Dict[str, Any]],
    text: str,
    model_name: str,
    api_key: Optional[str],
) -> List[str]:
    """
    2–3 sentence explanation per QA (same order).
    Requires API key (env/arg). If missing, raises; keep the call guarded in pipeline if you allow offline mode.
    """
    genai = configure_gemini(api_key)
    import google.generativeai as genai_pkg  # for typing clarity
    m = genai_pkg.GenerativeModel(model_name)

    # normalize QA tuples
    items: List[Tuple[str,str]] = []
    for it in qa:
        if isinstance(it, dict):
            q = str(it.get("question", "")).strip()
            a = str(it.get("answer", "")).strip()
        else:
            q, a = it  # type: ignore
        if q and a:
            items.append((q, a))
    pack = [{"q": q, "a": a} for q, a in items]

    prompt = (
        "For each item, produce a clear, in-depth explanation (4–6 sentences). "
        "Each explanation should:\n"
        "- Justify why the correct answer is right.\n"
        "- Contrast it with why alternative options (if any) are wrong.\n"
        "- Provide conceptual context, not just restatement.\n\n"
        "Return STRICT JSON: a list of strings in the same order; "
        "its length must equal the input list.\n\n"
        f"LECTURE (truncated):\n{text[:8000]}\n\n"
        f"ITEMS:\n{_json.dumps(pack, ensure_ascii=False)}"
    )

    resp = m.generate_content(prompt)
    data = getattr(resp, "text", "") or "[]"
    start, end = data.find("["), data.rfind("]")
    if start != -1 and end != -1:
        data = data[start:end+1]
    try:
        arr = _json.loads(data)
        out = [str(x) for x in arr]
        if len(out) != len(items):
            # length mismatch → minimal fallback (still useful)
            return [f"Explanation: {a[:200]}" for _, a in items]
        return out
    except Exception:
        return [f"Explanation: {a[:200]}" for _, a in items]


def infer_title(
    lecture_text: str,
    filename_stem: str,
    model_name: str,
    api_key: Optional[str],
) -> str:
    """
    Try a concise 3–7 word title via Gemini; else heuristic or file name.
    Requires API key (env/arg). If missing, returns heuristic.
    """
    key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=key)
            m = genai.GenerativeModel(model_name)
            prompt = (
                "Give a concise 3–7 word title for this lecture text. "
                "Return ONLY the title, no quotes, no trailing punctuation.\n\n"
                f"{lecture_text[:8000]}"
            )
            t = (m.generate_content(prompt).text or "").replace("\n", " ").strip(" .,:;\"'")
            if len(t.split()) >= 2:
                return t[:80]
        except Exception:
            pass

    # Heuristic fallback: first decent line
    for line in lecture_text.splitlines():
        s = re.sub(r"[^A-Za-z0-9 :/\-\(\)\[\]]+", "", line).strip()
        if len(s.split()) >= 2 and len(s) >= 10:
            return s[:80]
    return filename_stem[:80] or "Auto Quiz"
