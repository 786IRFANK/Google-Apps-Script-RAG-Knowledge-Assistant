import json
import re
import traceback
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from gemni_services import get_query_embedding, ask_gemini_or_plan, analyze_formulas
from retrieval_service import search, keyword_search
from session_service import session_service

app = FastAPI(title="TeamOne AI Backend", version="3.1")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


SOURCE_URL_MAP = {
    "formulas.md": {
        "title": "TeamOne | Formulas",
        "url"  : "https://succeedium.com/teamone/doc/formulas",
    },
    "functions.md": {
        "title": "TeamOne | Functions",
        "url"  : "https://succeedium.com/teamone/doc/functions",
    },
}

def filename_to_link(filename: str) -> dict:
    """Map a knowledge.json source filename to a {title, url} dict."""
    key = filename.rsplit("/", 1)[-1]          # strip any path prefix
    if key in SOURCE_URL_MAP:
        return SOURCE_URL_MAP[key]
    slug  = key.replace(".md", "").replace("_", "-").lower()
    label = key.replace(".md", "").replace("_", " ").title()
    return {
        "title": f"TeamOne | {label}",
        "url"  : f"https://succeedium.com/teamone/doc/{slug}",
    }

_SHEET_QUESTION_PATTERNS = [
    r"\bactive sheet\b",
    r"\bcurrent sheet\b",
    r"\bthis sheet\b",
    r"\bmy sheet\b",
    r"\bcurrent formula\b",
    r"\bthis formula\b",
    r"\banalyz(e|ing)\b.*\b(sheet|formula)\b",
    r"\breview(ing)?\b.*\b(sheet|formula)s?\b",
    r"\bsyntax error\b",
    r"\bargument issue",
    r"\bcompare\b.*\b(formula|documentation)\b",
    r"\b[A-Za-z][A-Za-z0-9 _]{0,20}![A-Z]{1,3}\d{1,5}\b",   # "Revenue!B8"
    r"\bcell\s+[A-Z]{1,3}\d{1,5}\b",                         # "cell B8"
    r"\b[A-Z]{1,2}\d{1,4}\b",                                # bare "B8" (loose, intentional)
]
_SHEET_QUESTION_RE = re.compile("|".join(_SHEET_QUESTION_PATTERNS), re.IGNORECASE)


def looks_like_sheet_question(question: str) -> bool:
    """True if the question is about the user's actual sheet/formula (needs
    sheet_context) rather than a pure documentation question."""
    return bool(_SHEET_QUESTION_RE.search(question or ""))


_EXPLICIT_SHEET_REFRESH_PATTERNS = [
    r"\bactive sheet\b",
    r"\bcurrent sheet\b",
    r"\bthis sheet\b",
    r"\bmy sheet\b",
]
_EXPLICIT_SHEET_REFRESH_RE = re.compile("|".join(_EXPLICIT_SHEET_REFRESH_PATTERNS), re.IGNORECASE)


def wants_fresh_sheet_fetch(question: str) -> bool:

    return bool(_EXPLICIT_SHEET_REFRESH_RE.search(question or ""))


def detect_function_name(question: str, sheet_context, last_cell: str = None) -> Optional[str]:
   
    if sheet_context is not None and last_cell:
        # last_cell may be "Sheet!B8" or bare "B8" — match on the bare cell.
        bare = last_cell.split("!")[-1]
        for f in sheet_context.formulas:
            if f.cell == bare and f.name in ("TEAMONE", "MDX"):
                return f.name

    text = (question or "").upper()
    for fn in ("TEAMONE", "MDX"):
        if fn in text:
            return fn

    if sheet_context is not None:
        for f in sheet_context.formulas:
            if f.name in ("TEAMONE", "MDX"):
                return f.name
    return None


def merge_contexts(*context_lists, limit=10):
    """Dedupe and cap combined retrieval results (semantic + keyword)."""
    seen = set()
    merged = []
    for lst in context_lists:
        for c in lst:
            key = (c["source"], c["text"][:120])
            if key in seen:
                continue
            seen.add(key)
            merged.append(c)
            if len(merged) >= limit:
                return merged
    return merged


def model_to_dict(model) -> dict:
    """pydantic v1/v2-safe conversion to a plain dict."""
    if hasattr(model, "model_dump"):
        return model.model_dump()
    return model.dict()


class SheetFormula(BaseModel):
    cell   : str        # e.g. "B8"
    formula: str        # JSON string — see docstring above
    name   : str = ""   # "TEAMONE" | "MDX" | "" if unknown

class SheetContext(BaseModel):
    
    sheetName: str
    formulas : list[SheetFormula] = []

class HistoryMessage(BaseModel):
    role   : str    # "user" | "assistant"
    content: str

class ChatRequest(BaseModel):
    question    : str
    session_id  : str                         = "default"   # see WARNING below
    history     : list[HistoryMessage]        = []          # full conversation from sidebar
    sheet_context: Optional[SheetContext]     = None        # None for general questions

class Source(BaseModel):
    title         : str
    url           : str
    score         : float = 0.0
    low_confidence: bool  = False

class FormulaPlan(BaseModel):
    cube   : str
    rows   : list[list[str]] = []
    columns: list[list[str]] = []
    titles : list[str]       = []
    options: str             = ""

class ChatResponse(BaseModel):
    intent      : str                     # "answer" | "clarify" | "formula_plan" | "need_sheet_context"
    message     : str
    formula_plan: Optional[FormulaPlan]   = None
    sources     : list[Source]

# Existing analyze models — unchanged
class GsInfoFormulaItem(BaseModel):
    cell   : str
    name   : str
    formula: str

class FormulaAnalyzeRequest(BaseModel):
    formulas: list[GsInfoFormulaItem]


# ---------------------------------------------------------------------------
# Helper: build deduplicated Source list with real URLs
# ---------------------------------------------------------------------------

def build_sources(contexts: list) -> list[dict]:
    seen, out = set(), []
    for c in contexts:
        link = filename_to_link(c["source"])
        url  = link["url"]
        if url in seen:
            continue
        seen.add(url)
        out.append({
            "title"         : link["title"],
            "url"           : url,
            "score"         : round(c["score"], 3),
            "low_confidence": bool(c.get("low_confidence", False)),
        })
    return out


# ---------------------------------------------------------------------------
# /chat  ── unified endpoint
# ---------------------------------------------------------------------------

@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    try:
        question = req.question.strip()
        print(question)
        if req.session_id == "default":
            # Every user hitting this default bucket means they all share ONE
            # conversation history and sheet cache with each other. This is
            # almost certainly a contributor to "inconsistent answers" /
            # "lost context" if the Apps Script sidebar isn't sending a real
            # per-user session_id. Fix on the sidebar side: generate a UUID
            # once and store it (e.g. PropertiesService), then send it as
            # session_id on every request.
            print("[chat] WARNING: session_id='default' — sessions are not "
                  "isolated per user. See comment in main.py.")


        effective_sheet_context = req.sheet_context
        sheet_context_is_cached = False
        force_fresh_fetch = req.sheet_context is None and wants_fresh_sheet_fetch(question)

        if effective_sheet_context is not None:
            session_service.set_sheet_context(req.session_id, model_to_dict(effective_sheet_context))
        elif not force_fresh_fetch:
            cached = session_service.get_last_sheet_context(req.session_id)
            if cached is not None:
                effective_sheet_context = SheetContext(**cached)
                sheet_context_is_cached = True


        if effective_sheet_context is None and (looks_like_sheet_question(question) or force_fresh_fetch):
            msg = (
                "I don't have your active sheet's formulas yet. Please click "
                "**Analyze active sheet** (or make sure the sheet you want "
                "checked is the active tab) so I can pull the current "
                "formulas, then ask again."
            )
            session_service.save(req.session_id, "user", question)
            session_service.save(req.session_id, "assistant", msg)
            return {
                "intent"      : "need_sheet_context",
                "message"     : msg,
                "formula_plan": None,
                "sources"     : [],
            }

        cell_ref = session_service.extract_cell_reference(
            question,
            effective_sheet_context.sheetName if effective_sheet_context else None,
        )
        if cell_ref:
            session_service.set_last_cell(req.session_id, cell_ref)
        last_cell = session_service.get_last_cell(req.session_id)

        embedding = get_query_embedding(question)
        semantic_contexts = search(embedding, top_k=5)

        detected_function = detect_function_name(question, effective_sheet_context, last_cell)
        function_contexts = keyword_search(detected_function, max_results=8) if detected_function else []

        contexts = merge_contexts(semantic_contexts, function_contexts, limit=10)

        if req.history:
            history = [{"role": m.role, "content": m.content} for m in req.history]
        else:
            history = session_service.get_history(req.session_id)

        try:
            result = ask_gemini_or_plan(
                question               = question,
                history                = history,
                contexts               = contexts,
                sheet_context          = effective_sheet_context,
                sheet_context_is_cached= sheet_context_is_cached,
                last_cell              = last_cell,
                detected_function      = detected_function,
            )
        except (json.JSONDecodeError, ValueError) as parse_err:
            print(f"[chat] WARNING Gemini response could not be parsed: {parse_err}")
            result = {
                "intent"      : "answer",
                "message"     : "Sorry, I ran into an issue processing that — could you rephrase or ask again?",
                "formula_plan": None,
            }

        intent       = result.get("intent", "answer")
        formula_plan = result.get("formula_plan")

        if intent == "formula_plan":
            valid = (
                isinstance(formula_plan, dict)
                and formula_plan.get("cube")
                and isinstance(formula_plan.get("rows", []), list)
                and isinstance(formula_plan.get("columns", []), list)
            )
            if not valid:
                print(f"[chat] WARNING formula_plan invalid: {formula_plan}")
                intent, formula_plan = "answer", None

        session_service.save(req.session_id, "user",      question)
        session_service.save(req.session_id, "assistant", result.get("message", ""))

        return {
            "intent"      : intent,
            "message"     : result.get("message", ""),
            "formula_plan": formula_plan,
            "sources"     : build_sources(contexts),
        }

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Chat failed: {e}")



@app.get("/health")
def health():
    return {"status": "ok", "version": "3.1"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=False)