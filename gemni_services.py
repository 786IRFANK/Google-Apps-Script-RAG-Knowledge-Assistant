import os
import re
import json
from google import genai
from google.genai import types
from google.genai.errors import APIError
from dotenv import load_dotenv
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

load_dotenv()

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

CHAT_MODEL  = os.getenv("GEMINI_API_KEY","gemini-2.5-flash")
EMBED_MODEL = os.getenv("GEMINI_API_KEY","gemini-embedding-001")

CHAT_TEMPERATURE = 0.1

api_retry_decorator = retry(
    stop    = stop_after_attempt(4),
    wait    = wait_exponential(multiplier=1, min=2, max=10),
    retry   = retry_if_exception_type(APIError),
    reraise = True,
)


@api_retry_decorator
def _call_embed_api(question):
    return client.models.embed_content(
        model    = EMBED_MODEL,
        contents = question,
        config   = types.EmbedContentConfig(task_type="RETRIEVAL_QUERY"),
    )


@api_retry_decorator
def _call_json_api(prompt, temperature=CHAT_TEMPERATURE):
    return client.models.generate_content(
        model    = CHAT_MODEL,
        contents = prompt,
        config   = types.GenerateContentConfig(
            response_mime_type = "application/json",
            temperature        = temperature,
        ),
    )


def get_query_embedding(question):
    response = _call_embed_api(question)
    if not response.embeddings:
        raise ValueError("Gemini returned no embedding for the query.")
    return response.embeddings[0].values


def _safe_json_parse(text: str) -> dict:
    """
    Gemini is instructed to return raw JSON, but models occasionally wrap it
    in ```json fences anyway. Previously any deviation raised a bare
    json.JSONDecodeError that bubbled up as a 500 and looked like "the chatbot
    isn't working". Strip fences defensively before parsing.
    """
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(json)?\s*|\s*```$", "", cleaned.strip(), flags=re.IGNORECASE)
    return json.loads(cleaned)


def _context_block(contexts: list) -> str:
    parts = []
    for c in contexts:
        flag = " [LOW CONFIDENCE]" if c.get("low_confidence") else ""
        parts.append(f"Source{flag}: {c['source']}\n\n{c['text']}")
    return "\n\n---\n\n".join(parts) if parts else "(No documentation retrieved.)"


def _history_block(history: list) -> str:
    return "\n".join(f"{h['role'].upper()}: {h['content']}" for h in history) or "(No prior conversation.)"



_ARG_POSITION_NOTE = (
    "Argument POSITIONS for both TEAMONE() and MDX() are always, in order: "
    "1) cube, 2) rows, 3) columns, 4) titles / context filters, "
    "5) options (flags), and for TEAMONE() only, 6) solveOrder."
)


def _describe_formula(f) -> str:
    """
    Render one SheetFormula into an unambiguous, pre-parsed description for
    the prompt. f.formula is the JSON string described above; f.name is the
    function (TEAMONE/MDX) if known.
    """
    try:
        parsed = json.loads(f.formula)
    except (json.JSONDecodeError, TypeError):
        # Not JSON — show as-is rather than guessing structure.
        return f"  {f.cell:<12} [{f.name or 'unknown function'}]  (could not parse) raw value: {f.formula}"

    options = str(parsed.get("options", "") or "")
    flags = list(options) if options else []

    lines = [
        f"  Cell: {f.cell}   Function: {f.name or '(unknown)'}",
        f"    Arg 1 (cube)              : {parsed.get('cube', '')!r}",
        f"    Arg 2 (rows)              : {parsed.get('rows', '')!r}",
        f"    Arg 3 (columns)           : {parsed.get('columns', '')!r}",
        f"    Arg 4 (titles/context)    : {parsed.get('titles', '')!r}",
        f"    Arg 5 (options/flags)     : {options!r}  -> individual flags: {flags or '(none)'}",
    ]
    if f.name == "TEAMONE":
        lines.append(f"    Arg 6 (solveOrder)        : {parsed.get('solveOrder', 0)!r}")
    return "\n".join(lines)


def _sheet_context_block(sheet_context, is_cached: bool) -> str:
    """
    Format the SheetContext into a clear, ALREADY-PARSED text block for the
    prompt (see note above on why raw formula text is never shown as-is).

    is_cached=True means this snapshot was NOT sent fresh on this turn — it's
    reused from earlier in the conversation, so follow-ups still work without
    the sidebar re-fetching every turn. The model is told it may be stale so
    it doesn't assert facts with false confidence if the user has since
    edited the cell or switched sheets.
    """
    if sheet_context is None:
        return "(Not provided — user asked a general documentation question, no sheet fetch needed.)"

    staleness_note = (
        " NOTE: this snapshot is being REUSED from earlier in the conversation "
        "(not re-fetched this turn). If the user seems to be describing a "
        "different formula or a different sheet than what's shown here, say so "
        "instead of guessing."
        if is_cached else ""
    )

    lines = [
        f"Active sheet name : {sheet_context.sheetName}{staleness_note}",
        f"Total formulas    : {len(sheet_context.formulas)}",
        "",
        _ARG_POSITION_NOTE,
        "",
    ]
    if not sheet_context.formulas:
        lines.append("No TEAMONE()/MDX() formulas found on this sheet.")
    else:
        for f in sheet_context.formulas:
            lines.append(_describe_formula(f))
            lines.append("")
    return "\n".join(lines)


# ===========================================================================
# Main chat function
# ===========================================================================

def ask_gemini_or_plan(
    question: str,
    history: list,
    contexts: list,
    sheet_context=None,
    sheet_context_is_cached: bool = False,
    last_cell: str = None,
    detected_function: str = None,
) -> dict:
    """
    Single entry point for ALL chat messages, including formula analysis and
    "analyze active sheet" — there is deliberately only one endpoint now, so
    a formula that was just analyzed and a follow-up question about it share
    the same conversation history and the same sheet-context snapshot.

    Returns a dict:
      {
        "intent"      : "answer" | "clarify" | "formula_plan",
        "message"     : str,
        "formula_plan": dict | None
      }
    """

    prompt = f"""
You are TeamOne AI Assistant, embedded in a Google Sheets sidebar.
You help users understand TeamOne documentation AND analyze formulas on their active sheet.

============================
YOUR TWO JOBS
============================

JOB 1 — Answer documentation questions, and analyze/validate formulas on the active sheet.
Use ONLY the DOCUMENTATION CONTEXT section below for what's valid/documented.
Use ONLY the ACTIVE SHEET CONTEXT section below (already parsed into its real
arguments — see rules) for what the user's actual formulas contain.
If the documentation does not cover something, say so clearly rather than guessing.

JOB 2 — Generate a TEAMONE()/MDX() formula plan.
When the user asks you to CREATE or GENERATE a formula layout:
- Do NOT copy-paste a formula from the docs.
- Build a structured plan using the exact values the user provided.
- If the user gave specific cube/rows/columns/filters → use those exact values.
- If no specifics were given → output intent "clarify" and ask what's missing.
- If user explicitly asks for a "generic/simple example" → use placeholder values
  and say so in "message".

============================
ACTIVE SHEET ANALYSIS — CRITICAL RULES
============================

When ACTIVE SHEET CONTEXT below contains formulas, these rules ALWAYS apply:

RULE 1 — THE FORMULA IS ALREADY PARSED. USE IT DIRECTLY, NEVER RE-PARSE.
  Each formula in ACTIVE SHEET CONTEXT is pre-split into its real arguments
  (cube, rows, columns, titles, options, solveOrder) by the spreadsheet
  itself. This is NOT formula text with commas to split — it is the true
  argument values. Read them directly from the labeled "Arg N (...)" lines.
  If the user asks about a cell not listed there, say: "I can see [N]
  formulas on sheet [name] but [cell] is not among them."

RULE 2 — COMPARE FLAGS AGAINST DOCUMENTATION, COMPLETELY AND CORRECTLY.
  The "options" argument (Arg 5) is the flags argument. Its "individual
  flags" list already shows each flag character separately — e.g. options
  "SI" means the flags present are S and I, nothing else.
  When asked to analyze flags or "what are the available options?":
  a) Take the flags list EXACTLY as given for Arg 5 — do not reinterpret,
     re-split, or invent flags that aren't in that list.
  b) List EVERY documented flag for that function from DOCUMENTATION CONTEXT
     — not just the ones present in the user's formula. If the user asks
     "why didn't you suggest the other options?" it means a previous answer
     was incomplete; enumerate ALL documented flags this time.
  c) For each flag actually present in Arg 5: mark VALID ✓ or INVALID ✗ based
     on whether the documentation lists it — do not mark invalid just because
     you are unsure, only when the docs explicitly do not list it.

RULE 3 — EXPLAIN ARGUMENTS BY POSITION, USING THE FIXED POSITIONS ABOVE.
  {_ARG_POSITION_NOTE}
  When asked "what is the second argument?" (or similar) about a specific
  cell, that is always the ROWS argument (Arg 2) — read its value directly
  from ACTIVE SHEET CONTEXT for that cell. Never derive position by counting
  commas in raw text; there is no raw text here, only the parsed arguments.

RULE 4 — ALWAYS INCLUDE SHEET NAME IN CELL REFERENCES.
  Always write cells as SheetName!CellRef (e.g. Revenue!B8, not just B8) in
  your response text, using the "Active sheet name" given above.

RULE 5 — USE CONVERSATION HISTORY AND SESSION MEMORY FOR CONTEXT.
  If the user refers to "the formula", "that cell", "the second argument" etc.
  without specifying which cell, use LAST DISCUSSED CELL / FUNCTION IN FOCUS
  below (already resolved for you), falling back to CONVERSATION HISTORY only
  if those are empty.

RULE 6 — NEVER HALLUCINATE SHEET CONTENT.
  If no ACTIVE SHEET CONTEXT is provided below, do not claim to know what is in
  any cell. Say: "I don't have your sheet data for this question yet."

RULE 7 — BE CONSISTENT.
  Base every claim strictly on DOCUMENTATION CONTEXT and ACTIVE SHEET CONTEXT
  below, not on prior turns' phrasing. If asked "are you sure?" or "why did
  your answer change?", re-derive the answer from the same context rather than
  defending a previous phrasing — if the docs support a more complete answer
  than before, give that, and briefly note what was incomplete previously.

============================
FORMULA PLAN STRUCTURE (for intent = "formula_plan")
============================
Match the gsTeamoneFormulaAdd(cube, rangeRows, rangeCols, rangeTitles, options) signature:

  cube    : string — the cube name
  rows    : array of arrays — one inner array per ROW, listing hierarchy labels
            e.g. [["Retail","Phones"],["Retail","PCs"]]
  columns : array of arrays — TRANSPOSED — one inner array per HEADER LEVEL
            e.g. [["Actual","Budget"],["Gross Revenue","Gross Revenue"]]
  titles  : array of "Dimension^Element" strings for context filters
            e.g. ["Year^2026","Version^Actual"]
  options : string of option flags e.g. "S" or ""

============================
SESSION MEMORY (already resolved server-side — trust this over guessing)
============================
Last cell explicitly discussed : {last_cell or "(none yet this session)"}
Function currently in focus    : {detected_function or "(not detected)"}

============================
CONVERSATION HISTORY
============================
{_history_block(history)}

============================
DOCUMENTATION CONTEXT
============================
{_context_block(contexts)}

============================
ACTIVE SHEET CONTEXT (already parsed — see Rule 1)
============================
{_sheet_context_block(sheet_context, sheet_context_is_cached)}

============================
CURRENT QUESTION
============================
{question}

============================
OUTPUT FORMAT — STRICT
============================
Respond with ONLY a valid JSON object. No markdown fences, no extra commentary.

{{
  "intent": "answer" | "clarify" | "formula_plan",
  "message": "<what to show the user in the chat — Markdown OK>",
  "formula_plan": {{
    "cube"   : "...",
    "rows"   : [[...]],
    "columns": [[...]],
    "titles" : [...],
    "options": "..."
  }} or null
}}

For intent "answer"  → formula_plan must be null.
For intent "clarify" → formula_plan must be null.
For intent "formula_plan" → formula_plan must be a valid object with at least cube set.
"""

    response = _call_json_api(prompt)
    return _safe_json_parse(response.text)


def analyze_formulas(items: list) -> str:
    """
    Audit pre-parsed TEAMONE/MDX formula arguments against documentation.
    items: list of {cell, name, formula (JSON string), contexts}
    Returns a JSON array string.
    """
    blocks = []
    for item in items:
        name = item.get("name", "")
        if name not in ("TEAMONE", "MDX"):
            continue
        try:
            parsed = json.loads(item["formula"])
        except Exception:
            parsed = {"_raw": item.get("formula")}

        ctx_lines = []
        for c in item.get("contexts", []):
            snippet = c["text"][:500].replace("\n", " ")
            ctx_lines.append(f"- ({c['source']}) {snippet}")
        ctx_text = "\n".join(ctx_lines) if ctx_lines else "No relevant documentation found."

        blocks.append(
            f"Cell: {item['cell']}\n"
            f"Function: {name}\n"
            f"Parsed arguments: {json.dumps(parsed)}\n"
            f"Relevant Docs:\n{ctx_text}"
        )

    if not blocks:
        return "[]"

    prompt = f"""
You are a TeamOne formula auditor. For each item below you are given the ALREADY
PARSED arguments of a TEAMONE() or MDX() formula (cube, rows, columns, titles,
options), not the raw formula text. Using ONLY the provided documentation excerpts,
judge whether the arguments are structurally valid: correct types, required arguments
present, valid option flags per the docs, cube name present, at least one of
rows/columns/titles present.

Do NOT verify whether the cube or elements exist in TM1 — only check structural syntax.

{chr(10).join(blocks)}

Respond with ONLY a JSON array (no markdown fences), exactly:
[
  {{
    "cell"      : "D8",
    "valid"     : true,
    "issue"     : "",
    "suggestion": ""
  }}
]
Use valid=false and fill "issue" (and "suggestion" if fixable) only when the docs
clearly indicate a structural problem.
"""
    return _call_json_api(prompt).text