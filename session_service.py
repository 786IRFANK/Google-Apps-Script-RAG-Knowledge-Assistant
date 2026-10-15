import os
import re
import time
from collections import defaultdict


MAX_HISTORY = int(os.getenv("MAX_HISTORY", "60"))  # ~30 user/assistant turns

_CELL_REF_PATTERN = re.compile(r"\b([A-Za-z][A-Za-z0-9 _]{0,20}!)?([A-Z]{1,3}\d{1,5})\b")


class SessionService:

    def __init__(self):
        self.sessions = defaultdict(lambda: {
            "history": [],
            "last_sheet_context": None,
            "last_cell": None,
            "updated_at": time.time(),
        })

    # ---- history -----------------------------------------------------
    def get_history(self, session_id):
        return self.sessions[session_id]["history"]

    def save(self, session_id, role, content):
        hist = self.sessions[session_id]["history"]
        hist.append({"role": role, "content": content})
        if len(hist) > MAX_HISTORY:
            self.sessions[session_id]["history"] = hist[-MAX_HISTORY:]
        self.sessions[session_id]["updated_at"] = time.time()

    # ---- sheet context cache ------------------------------------------
    def get_last_sheet_context(self, session_id):
        return self.sessions[session_id]["last_sheet_context"]

    def set_sheet_context(self, session_id, sheet_context_dict):
        if sheet_context_dict is not None:
            self.sessions[session_id]["last_sheet_context"] = sheet_context_dict
            self.sessions[session_id]["updated_at"] = time.time()

    # ---- last-discussed cell ------------------------------------------
    def get_last_cell(self, session_id):
        return self.sessions[session_id]["last_cell"]

    def set_last_cell(self, session_id, cell_ref):
        if cell_ref:
            self.sessions[session_id]["last_cell"] = cell_ref

    def extract_cell_reference(self, text, sheet_name=None):
        """
        Pull the most likely cell reference out of free text.
        'check B8' -> 'B8' (or 'SheetName!B8' if sheet_name is known).
        If multiple cells are mentioned, the LAST one wins (most recent intent).
        """
        matches = _CELL_REF_PATTERN.findall(text or "")
        if not matches:
            return None
        sheet_prefix, cell = matches[-1]
        if sheet_prefix:
            return f"{sheet_prefix}{cell}"
        if sheet_name:
            return f"{sheet_name}!{cell}"
        return cell


session_service = SessionService()