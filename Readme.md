# Google Apps Script + RAG Knowledge Assistant

Connects a retrieval-augmented generation (RAG) chatbot, backed by TeamOne documentation, to a Google Sheets sidebar. Users ask questions about spreadsheet formulas; Apps Script calls a FastAPI endpoint and displays the response.

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Backend Setup (FastAPI)](#backend-setup-fastapi)
- [Building the Knowledge Base](#building-the-knowledge-base)
- [Frontend Setup (Apps Script)](#frontend-setup-apps-script)
- [API Reference](#api-reference)
- [Sheet Context](#sheet-context)
- [Analysis Output Format](#analysis-output-format)
- [Configuration](#configuration)
- [Testing](#testing)
- [Troubleshooting](#troubleshooting)
- [Security Notes](#security-notes)

---

## Overview

The assistant lives in the TeamOne Google Sheets sidebar and helps users:

- Ask questions about TeamOne documentation (`TEAMONE()`, `MDX()`, `ATTR()`, `SUBSET()`).
- Analyze and validate formulas on the active sheet or in a specific cell.
- Compare what a formula does against what the documentation says it should do.
- Follow up across turns, with the full conversation history and the last known sheet context retained per session.

Answers include deduplicated, clickable links to the relevant section of the formulas documentation.

## Architecture

```
┌────────────────────────────────────────────┐
│ Google Sheets                              │
│  ┌──────────────────────────────────────┐  │
│  │ Sidebar (Vue 2 + Vuetify)            │  │
│  │  core.html / core_js.html            │  │
│  └───────────────┬──────────────────────┘  │
│                  │ google.script.run       │
│  ┌───────────────▼──────────────────────┐  │
│  │ Apps Script (teamone.js)             │  │
│  │  gsGetAssistantSheetContext()        │  │
│  └───────────────┬──────────────────────┘  │
└──────────────────┼─────────────────────────┘
                   │ fetch() POST /chat
                   │ {session_id, message, sheet_context}
         ┌─────────▼──────────┐
         │ FastAPI backend    │
         │  • intent classify │
         │  • RAG retrieval   │──► knowledge.json (embedded doc chunks)
         │  • Gemini answer   │──► Gemini API
         └────────────────────┘
```

Flow for each message:

1. The sidebar calls Apps Script to collect the active sheet name, active cell, the real formula text, and every `TEAMONE()` / `MDX()` formula on the sheet.
2. The sidebar sends the question plus this `sheet_context` directly to the FastAPI `/chat` endpoint using `fetch()` (not proxied through Apps Script, for speed).
3. The backend classifies the intent, retrieves relevant documentation chunks, and generates an answer with Gemini using the full session history.
4. The sidebar renders the answer and the source links.

## Project Structure

```
.
├── backend/
│   ├── main.py              # FastAPI app: /chat endpoint, intent routing, RAG, prompts
│   ├── build_knowledge.py   # Chunks docs and builds embeddings (knowledge.json)
│   ├── knowledge.json       # Embedded documentation chunks
│   ├── docs/
│   │   ├── formulas.md
│   │   └── functions.md
│   ├── requirements.txt
│   └── .env
└── apps-script/
    ├── teamone.js           # Server-side Apps Script (sheet context helpers)
    ├── core.html            # Sidebar markup (Vue 2 + Vuetify)
    └── core_js.html         # Sidebar logic (chat, fetch, sheet context)
```

> Adjust file names to match your repository layout.

## Prerequisites

- Python 3.10+
- A Google Gemini API key
- A Google account with access to the TeamOne Sheets add-on project (Apps Script)
- A publicly reachable HTTPS URL for the backend (Apps Script sidebars call it from the browser)

## Backend Setup (FastAPI)

```bash
cd backend
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

Create `backend/.env`:

```env
GEMINI_API_KEY=your_gemini_api_key
CHAT_MODEL=gemini-2.5-flash
EMBEDDING_MODEL=gemini-embedding-001
KNOWLEDGE_FILE=knowledge.json
KNOWLEDGE_SOURCE_FILTER=formulas.md
LOG_LEVEL=INFO
```

Run locally:

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Interactive docs are available at `http://localhost:8000/docs`.

## Building the Knowledge Base

The knowledge base is built from the TeamOne documentation markdown files (`formulas.md`, `functions.md`). Each file is chunked and embedded with `gemini-embedding-001`, then saved to `knowledge.json`.

```bash
python build_knowledge.py
```

Re-run this whenever the documentation changes, then restart the backend.

## Frontend Setup (Apps Script)

1. Open the Google Sheet and go to **Extensions → Apps Script**.
2. Add or update the files:
   - `teamone.js` — includes `gsGetAssistantSheetContext()`
   - `core.html` — sidebar UI
   - `core_js.html` — sidebar logic, including the `gsRunAsync()` Promise wrapper around `google.script.run`
3. Set the backend base URL in `core_js.html` (for example `const ASSISTANT_API_URL = "https://your-backend.example.com";`).
4. Save and reload the sheet, then open the TeamOne sidebar.

## API Reference

### `POST /chat`

**Request**

```json
{
  "session_id": "abc-123",
  "message": "Review active sheet formulas",
  "sheet_context": {
    "sheet_name": "Budget",
    "active_cell": "B2",
    "formula": "=TEAMONE(...)",
    "formulas": [
      { "cell": "B2", "sheet_name": "Budget", "formula": "=TEAMONE(...)" }
    ]
  }
}
```

**Response**

```json
{
  "session_id": "abc-123",
  "intent": {
    "intent": "active_sheet_analysis",
    "requires_sheet": true,
    "requires_documentation": true,
    "requires_history": true,
    "cell": null,
    "sheet_name": "Budget",
    "reason": "User asked to review formulas on the active sheet"
  },
  "message": "Sheet: Budget\n...",
  "sources": [
    { "label": "TEAMONE function", "url": "https://succeedium.com/teamone/doc/formulas#teamone-function" }
  ]
}
```

Notes:

- `session_id` keeps conversation history on the backend. The full history is used; there is no compaction or truncation.
- `sheet_context` may be `null` for pure documentation questions. If omitted on a follow-up turn, the backend falls back to the last sheet context stored for the session.
- Formulas that reference documentation (analysis, validation, comparison intents) always trigger retrieval, regardless of the classifier output.

## Sheet Context

`gsGetAssistantSheetContext()` in `teamone.js` builds the `sheet_context` payload:

- Active sheet name and active cell.
- The real formula text via `getFormula()` / `getFormulas()`. (It does **not** use the `gsInfo(['FORMULAS'])` value, which is a JSON cache blob keyed to the last recalculated cell and unreliable for arbitrary cells.)
- A scan of the sheet's used range for every `TEAMONE()` and `MDX()` formula, returned in `formulas[]`.

The sidebar awaits this through `gsRunAsync()` and sends it with every chat message.

## Analysis Output Format

Analysis-type answers follow a fixed structure:

```
Sheet: <sheet name>

Found formulas:
- <cell>: FUNCTION()
- <cell>: FUNCTION()

<cell> — FUNCTION()
Issue: ...
Looks valid: ...
Suggestion: ...
```

Formulas with no problems are summarized in a one-line "looks valid" note instead of the full block. The format is defined once as a shared prompt template and used by the active-sheet, formula-analysis, cell-formula, formula-validation, and comparison answers.

## Configuration

| Variable | Description | Default |
|----------|-------------|---------|
| `GEMINI_API_KEY` | Gemini API key | — |
| `CHAT_MODEL` | Model used for classification and answers | `gemini-2.5-flash` |
| `EMBEDDING_MODEL` | Model used for document embeddings | `gemini-embedding-001` |
| `KNOWLEDGE_FILE` | Path to the embedded knowledge file | `knowledge.json` |
| `KNOWLEDGE_SOURCE_FILTER` | Restrict retrieval to a source file (only applies if chunks carry source metadata) | `formulas.md` |
| `LOG_LEVEL` | Logging verbosity | `INFO` |

## Testing

Test the backend directly first to separate backend problems from Apps Script problems.

```bash
curl -X POST http://localhost:8000/chat \
  -H "Content-Type: application/json" \
  -d '{
    "session_id": "test-1",
    "message": "What does the S flag do in TEAMONE()?",
    "sheet_context": null
  }'
```

With a sheet context:

```bash
curl -X POST http://localhost:8000/chat \
  -H "Content-Type: application/json" \
  -d '{
    "session_id": "test-2",
    "message": "Analyze cell B2",
    "sheet_context": {
      "sheet_name": "Budget",
      "active_cell": "B2",
      "formula": "=TEAMONE(\"server\",\"cube\",\"S\")",
      "formulas": [
        {"cell": "B2", "sheet_name": "Budget", "formula": "=TEAMONE(\"server\",\"cube\",\"S\")"}
      ]
    }
  }'
```

The same requests can be sent from Postman.

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| "No formulas found" | `sheet_context` is `null` or empty | Confirm `gsGetAssistantSheetContext()` runs and the fetch body includes `sheet_context` |
| "No documentation provided" | Retrieval returned nothing | Check `knowledge.json` exists and was rebuilt; confirm function names are detected in the query |
| Source links not clickable | Old `core.html` chips | Ensure chips render `href`/`target` from `result.sources` |
| CORS error in the sidebar | Backend missing CORS config | Allow the Apps Script sandbox origins in FastAPI `CORSMiddleware` |
| Slow or timed-out responses | Large sheet or cold backend | Limit scanned range, check Gemini latency, use a warm instance |
| Follow-ups lose sheet context | New `session_id` each message | Reuse one `session_id` per sidebar session |
| Backend works in curl but not in the sidebar | Apps Script side issue | Check the browser console and the Apps Script execution log |

## Security Notes

- Keep `GEMINI_API_KEY` in `.env` or a secrets manager; never commit it or ship it to the sidebar.
- Serve the backend over HTTPS only.
- Restrict CORS origins to what the sidebar needs.
- Sheet formulas may contain server or cube names; avoid logging full `sheet_context` in production.
