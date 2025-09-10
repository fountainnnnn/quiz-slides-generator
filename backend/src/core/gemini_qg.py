from typing import Any, Dict, List, Tuple, Optional
import json as _json

def configure_gemini(api_key: Optional[str] = None):
    import google.generativeai as genai
    # Embedded default key (you can override via the UI or env vars)
    key = (api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
           or "AIzaSyBUHRWVLSwPfbAcSTyueFUX7PHQzNKPd4c")
    if not key:
        raise RuntimeError("Gemini API key not found. Provide via UI or set GEMINI_API_KEY/GOOGLE_API_KEY env var.")
    genai.configure(api_key=key)
    return genai

GEMINI_DEFAULT_MODEL = "gemini-1.5-flash"  # fast & capable

QG_SYSTEM = (
    "You are an expert quiz generator for lecturers. "
    "Given slide text in tidy markdown, generate rigorous, unambiguous questions grounded ONLY in the provided content. "
    "Allowed types: 'mcq', 'theory', 'code_fill'.\n"
    "- 'mcq': provide 4 options; 'answer' MUST be the exact option text; include a brief 'explanation'.\n"
    "- 'theory': short-answer or conceptual question; put the ideal answer in 'answer'; include 'explanation'.\n"
    "- 'code_fill': provide a prompt + code block with blanks like `___`; put the exact filled line(s) in 'answer'; brief 'explanation'.\n"
    "Always include 'source_slide_index' starting from 1. "
    "Return STRICT JSON: a list of objects with keys "
    "['type','question','options','answer','explanation','source_slide_index']. No prose."
)

def chunk_slides_for_qg(slide_md_paths: List[str], max_chars_per_chunk: int = 8000):
    chunks = []; buf = []; buf_len = 0; idxs = []
    for i, p in enumerate(slide_md_paths, 1):
        t = Path(p).read_text(encoding="utf-8")
        block = f"\n\n<!-- SLIDE {i} -->\n{t}\n"
        if buf_len + len(block) > max_chars_per_chunk and buf:
            chunks.append((idxs, "".join(buf))); buf = []; buf_len = 0; idxs = []
        buf.append(block); buf_len += len(block); idxs.append(i)
    if buf: chunks.append((idxs, "".join(buf)))
    return chunks

def build_qg_prompt(slide_block: str, want_counts: Dict[str,int], difficulty: str):
    total = sum(want_counts.values())
    mix_desc = ", ".join([f"{k}:{v}" for k,v in want_counts.items() if v>0]) or "auto"
    return f"""{QG_SYSTEM}

Difficulty target: {difficulty}.
Desired total questions in this call: {total} with mix {mix_desc}. 
If there isn't enough material for some type, reallocate to others.

Slide content:
{slide_block}

Return JSON only.
"""

def safe_json_parse(s: str) -> List[Dict[str, Any]]:
    try:
        return _json.loads(s)
    except Exception:
        m = re.search(r'\[.*\]', s, re.S)
        if m:
            try: return _json.loads(m.group(0))
            except Exception: pass
    return []

def generate_questions_with_gemini(
    per_slide_md_paths: List[str],
    total_questions: int = 20,
    mix: str = "auto",  # "auto" or "balanced" or "custom"
    custom_counts: Optional[Dict[str,int]] = None,
    difficulty: str = "mixed",
    model_name: str = GEMINI_DEFAULT_MODEL,
    api_key: Optional[str] = None
) -> List[Dict[str, Any]]:
    genai = configure_gemini(api_key)
    model = genai.GenerativeModel(model_name)

    if mix == "custom" and custom_counts:
        desired = {"mcq": int(custom_counts.get("mcq",0)), "theory": int(custom_counts.get("theory",0)), "code_fill": int(custom_counts.get("code_fill",0))}
        if sum(desired.values()) <= 0: desired = {"mcq": total_questions}
    elif mix == "balanced":
        per = max(1, total_questions // 3)
        desired = {"mcq": per, "theory": per, "code_fill": total_questions - 2*per}
    else:
        mcq = math.ceil(total_questions * 0.5)
        theory = max(0, math.ceil(total_questions * 0.3))
        code_fill = max(0, total_questions - mcq - theory)
        desired = {"mcq": mcq, "theory": theory, "code_fill": code_fill}

    chunks = chunk_slides_for_qg(per_slide_md_paths, max_chars_per_chunk=8000)
    remaining = total_questions
    results: List[Dict[str, Any]] = []

    for idxs, block in chunks:
        if remaining <= 0: break
        per_chunk = max(1, min(remaining, math.ceil(total_questions * (len(idxs) / len(per_slide_md_paths)))))

        want_counts = desired.copy()
        s = sum(want_counts.values())
        if s > 0:
            for k in want_counts: want_counts[k] = max(0, round(want_counts[k] * per_chunk / s))
            diff = per_chunk - sum(want_counts.values())
            for k in ["mcq","theory","code_fill"]:
                if diff == 0: break
                want_counts[k] += 1; diff -= 1

        prompt = build_qg_prompt(block, want_counts, difficulty)
        resp = model.generate_content(prompt)
        text = getattr(resp, "text", None) or (resp.candidates[0].content.parts[0].text if getattr(resp, "candidates", None) else "")
        got = safe_json_parse(text)

        clean = []
        for q in got:
            if not isinstance(q, dict): continue
            t = q.get("type","").lower()
            if t not in {"mcq","theory","code_fill"}: continue
            question = (q.get("question") or "").strip()
            answer = (q.get("answer") or "").strip()
            if not question or not answer: continue
            opts = q.get("options", [])
            if t == "mcq":
                if not isinstance(opts, list) or len(opts) < 3: continue
                opts = [str(o).strip() for o in opts if str(o).strip()]
                if answer not in opts:
                    if answer.upper() in ["A","B","C","D","E"] and len(opts) >= ord(answer.upper())-65+1:
                        answer = opts[ord(answer.upper())-65]
                    else:
                        continue
            else:
                opts = []
            exp = (q.get("explanation") or "").strip()
            src = int(q.get("source_slide_index") or idxs[0])
            clean.append({"type": t, "question": question, "options": opts, "answer": answer, "explanation": exp, "source_slide_index": src})

        clean = clean[:per_chunk]
        results.extend(clean)
        remaining -= len(clean)

    return results
