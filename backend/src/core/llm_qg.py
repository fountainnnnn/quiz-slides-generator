# backend/src/core/openai_qg.py

from typing import Any, Dict, List, Tuple, Optional
from pathlib import Path
import os, math, re, json as _json
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import OpenAI

DEEPSEEK_DEFAULT_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-flash")

# ------------------------------------------------------------
# OpenAI configuration (env only — no embedded default key)
# ------------------------------------------------------------
def configure_deepseek(api_key: Optional[str] = None) -> OpenAI:
    """
    Configure OpenAI client with a required API key.

    Precedence:
      1) function arg `api_key`
      2) env var DEEPSEEK_API_KEY

    If none is set, raises a clear error.
    """
    key = api_key or os.environ.get("DEEPSEEK_API_KEY")
    if not key:
        raise RuntimeError(
            "DeepSeek API key not found. Set DEEPSEEK_API_KEY on the backend "
            "or pass api_key in the request."
        )
    return OpenAI(
        api_key=key,
        base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
    )


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
        buf.append(block)
        buf_len += len(block)
        idxs.append(i)
    if buf:
        chunks.append((idxs, "".join(buf)))
    return chunks


def build_qg_prompt(slide_block: str, want_counts: Dict[str, int], difficulty: str):
    total = max(1, sum(max(0, v) for v in want_counts.values()))
    mix_desc = ", ".join([f"{k}:{v}" for k, v in want_counts.items() if v > 0]) or "auto"
    return f"""{QG_SYSTEM}

Difficulty target: {difficulty}.
Generate exactly {total} items with mix {mix_desc}. If the content cannot support a type, reallocate to others.

Slide content:
{slide_block}

Return JSON only as {{"questions":[...]}}."""


# ------------------------------------------------------------
# Parsing & validation
# ------------------------------------------------------------
def safe_json_parse(s: str) -> List[Dict[str, Any]]:
    """Parse JSON robustly (handles ```json fences, leading/trailing noise)."""
    if not s:
        return []
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s.strip(), flags=re.I | re.M)
    try:
        obj = _json.loads(s)
        if isinstance(obj, list):
            return obj
        if isinstance(obj, dict) and isinstance(obj.get("questions"), list):
            return obj["questions"]
    except Exception:
        pass
    m = re.search(r"\[[\s\S]*\]", s)
    if m:
        try:
            obj = _json.loads(m.group(0))
            return obj if isinstance(obj, list) else []
        except Exception:
            return []
    return []


def _count_types(items: List[Dict[str, Any]]) -> Dict[str, int]:
    counts = {"mcq": 0, "theory": 0, "code_fill": 0, "fill_blank": 0}
    for item in items:
        t = str(item.get("type", "")).lower().strip()
        if t in counts:
            counts[t] += 1
    return counts


def _deficit_counts(desired: Dict[str, int], current: Dict[str, int], remaining: int) -> Dict[str, int]:
    deficits = {k: max(0, int(desired.get(k, 0)) - int(current.get(k, 0))) for k in desired}
    if sum(deficits.values()) == 0:
        deficits = {"mcq": remaining, "theory": 0, "code_fill": 0, "fill_blank": 0}
    while sum(deficits.values()) > remaining:
        for key in ["mcq", "theory", "code_fill", "fill_blank"]:
            if sum(deficits.values()) <= remaining:
                break
            if deficits.get(key, 0) > 0:
                deficits[key] -= 1
    while sum(deficits.values()) < remaining:
        deficits["theory"] = deficits.get("theory", 0) + 1
    return deficits


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

        question = (raw.get("question") or raw.get("text_with_blanks") or "").strip()
        answer = _coerce_answer_to_str(raw.get("answer"))
        if not question or not answer:
            continue

        opts: List[str] = []
        code = ""
        if t == "mcq":
            opts_raw = raw.get("options", [])
            if not isinstance(opts_raw, list):
                continue
            opts = [str(o).strip() for o in opts_raw if str(o).strip()]
            if len(opts) < 3:
                continue
            if answer not in opts and answer.upper() in ["A", "B", "C", "D", "E"]:
                idx = ord(answer.upper()) - 65
                if 0 <= idx < len(opts):
                    answer = opts[idx]
            if answer not in opts:
                continue
        elif t == "code_fill":
            code = (raw.get("code") or "").strip()
            opts = [str(o).strip() for o in (raw.get("options") or []) if str(o).strip()]
        else:
            opts = []

        exp = (raw.get("explanation") or "").strip()
        try:
            src = int(raw.get("source_slide_index", 0)) or (idxs_fallback[0] if idxs_fallback else 1)
        except Exception:
            src = idxs_fallback[0] if idxs_fallback else 1

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
# Public helpers
# ------------------------------------------------------------
def generate_qa(
    per_slide_md_paths: List[str],
    total_questions: int = 20,
    mix: str = "auto",
    custom_counts: Optional[Dict[str, int]] = None,
    difficulty: str = "mixed",
    model_name: str = DEEPSEEK_DEFAULT_MODEL,
    api_key: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Generate questions using OpenAI (key from env or arg). Strict JSON parsing with validation.
    Supports types: mcq, theory, code_fill, fill_blank.
    """
    client = configure_deepseek(api_key)

    if mix == "custom" and custom_counts:
        desired = {
            "mcq": max(0, int(custom_counts.get("mcq", 0))),
            "theory": max(0, int(custom_counts.get("theory", 0))),
            "code_fill": max(0, int(custom_counts.get("code_fill", 0))),
            "fill_blank": max(0, int(custom_counts.get("fill_blank", 0))),
        }
        if sum(desired.values()) <= 0:
            desired = {"mcq": total_questions}
    elif mix == "balanced":
        types = ["mcq", "theory", "code_fill", "fill_blank"]
        base, remainder = divmod(max(1, total_questions), len(types))
        desired = {kind: base + (1 if idx < remainder else 0) for idx, kind in enumerate(types)}
    else:
        mcq = math.ceil(total_questions * 0.45)
        theory = max(0, math.ceil(total_questions * 0.25))
        code_fill = max(0, math.ceil(total_questions * 0.15))
        fill_blank = max(0, total_questions - mcq - theory - code_fill)
        desired = {"mcq": mcq, "theory": theory, "code_fill": code_fill, "fill_blank": fill_blank}

    chunks = chunk_slides_for_qg(per_slide_md_paths, max_chars_per_chunk=8000)
    results: List[Dict[str, Any]] = []
    focused_batch_size = max(1, min(int(os.getenv("QUIZ_GENERATOR_BATCH_SIZE", "5")), total_questions))
    max_parallel_calls = max(1, int(os.getenv("QUIZ_GENERATOR_MAX_PARALLEL_CALLS", "4")))

    def request_chunk(idxs: List[int], block: str, want_counts: Dict[str, int]) -> List[Dict[str, Any]]:
        prompt = build_qg_prompt(block, want_counts, difficulty)
        resp = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "system", "content": QG_SYSTEM}, {"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            extra_body={"thinking": {"type": "disabled"}},
        )
        text = resp.choices[0].message.content or ""
        return _clean_and_validate(safe_json_parse(text), idxs)

    work_items: List[Tuple[List[int], str, Dict[str, int]]] = []
    for idxs, block in chunks:
        weight = max(1, len(idxs))
        chunk_total = max(
            1, min(total_questions, math.ceil(total_questions * (weight / max(1, len(per_slide_md_paths)))))
        )

        want_counts = desired.copy()
        s = sum(want_counts.values()) or 1
        for k in want_counts:
            want_counts[k] = max(0, round(want_counts[k] * chunk_total / s))
        drift = chunk_total - sum(want_counts.values())
        for k in ["mcq", "theory", "code_fill", "fill_blank"]:
            if drift == 0:
                break
            want_counts[k] += 1
            drift -= 1

        expanded: List[str] = []
        for key in ["mcq", "theory", "code_fill", "fill_blank"]:
            expanded.extend([key] * want_counts.get(key, 0))
        for start in range(0, len(expanded), focused_batch_size):
            batch_counts = {"mcq": 0, "theory": 0, "code_fill": 0, "fill_blank": 0}
            for key in expanded[start:start + focused_batch_size]:
                batch_counts[key] += 1
            work_items.append((idxs, block, batch_counts))

    ordered_results: List[Tuple[int, List[Dict[str, Any]]]] = []
    with ThreadPoolExecutor(max_workers=max_parallel_calls) as pool:
        futures = {
            pool.submit(request_chunk, idxs, block, counts): idx
            for idx, (idxs, block, counts) in enumerate(work_items)
        }
        for future in as_completed(futures):
            try:
                ordered_results.append((futures[future], future.result()))
            except Exception:
                continue
    for _, chunk_results in sorted(ordered_results, key=lambda item: item[0]):
        results.extend(chunk_results)

    results = results[:total_questions]
    attempts = 0
    while len(results) < total_questions and attempts < 2 and chunks:
        attempts += 1
        remaining = total_questions - len(results)
        counts = _deficit_counts(desired, _count_types(results), remaining)
        idxs, block = chunks[0]
        try:
            results.extend(request_chunk(idxs, block, counts)[:remaining])
        except Exception:
            break

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
    client = configure_deepseek(api_key)

    items: List[Tuple[str, str]] = []
    for it in qa:
        if isinstance(it, dict):
            q = str(it.get("question", "")).strip()
            a = str(it.get("answer", "")).strip()
        else:
            q, a = it
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

    resp = client.chat.completions.create(
        model=model_name or DEEPSEEK_DEFAULT_MODEL,
        messages=[{"role": "user", "content": prompt}],
        extra_body={"thinking": {"type": "disabled"}},
    )
    data = resp.choices[0].message.content or "[]"
    start, end = data.find("["), data.rfind("]")
    if start != -1 and end != -1:
        data = data[start : end + 1]
    try:
        arr = _json.loads(data)
        out = [str(x) for x in arr]
        if len(out) != len(items):
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
    Try a concise 3–7 word title via OpenAI; else heuristic or file name.
    Requires API key (env/arg). If missing, returns heuristic.
    """
    try:
        client = configure_deepseek(api_key)
        prompt = (
            "Give a concise 3–7 word title for this lecture text. "
            "Return ONLY the title, no quotes, no trailing punctuation.\n\n"
            f"{lecture_text[:8000]}"
        )
        resp = client.chat.completions.create(
            model=model_name or DEEPSEEK_DEFAULT_MODEL,
            messages=[{"role": "user", "content": prompt}],
            extra_body={"thinking": {"type": "disabled"}},
        )
        t = (resp.choices[0].message.content or "").replace("\n", " ").strip(" .,:;\"'")
        if len(t.split()) >= 2:
            return t[:80]
    except Exception:
        pass

    for line in lecture_text.splitlines():
        s = re.sub(r"[^A-Za-z0-9 :/\-\(\)\[\]]+", "", line).strip()
        if len(s.split()) >= 2 and len(s) >= 10:
            return s[:80]
    return filename_stem[:80] or "Auto Quiz"
