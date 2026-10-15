"""Word MCP server: drives the running Microsoft Word through its COM automation API.

The API surface comes from MSWORD.OLB (decoded into word-api.json on first use). All COM work
runs on one dedicated STA thread; tools keep COM round trips low (whole-story text
reads, tab-separated table conversion) so calls stay in the few-millisecond range.
"""

from __future__ import annotations

import asyncio
import base64
import concurrent.futures
import io
import json
import os
import queue
import re
import sys
import threading
import time
from typing import Any, Callable

import pythoncom
import pywintypes
import win32com.client
from mcp.server.fastmcp import FastMCP, Image
from PIL import Image as PILImage

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # launched with python -I, which leaves the script folder off sys.path
from docx_context import Context  # noqa: E402

API_JSON = os.path.join(HERE, "word-api.json")
CALL_TIMEOUT = float(os.environ.get("WORD_MCP_TIMEOUT", "120"))

# COM HRESULTs worth handling.
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
RPC_E_DISCONNECTED = -2147417848
RPC_S_SERVER_UNAVAILABLE = -2147023174
CO_E_OBJNOTCONNECTED = -2147220995
MK_E_UNAVAILABLE = -2147221021

# Word constants (from the decoded enums).
WD_COLLAPSE_END = 0
WD_SEPARATE_BY_TABS = 1
WD_ALERTS_NONE = 0
WD_FORMAT = {
    "docx": 16, "doc": 0, "dotx": 14, "docm": 13, "pdf": 17, "xps": 18,
    "rtf": 6, "txt": 2, "html": 8, "htm": 8, "mht": 9, "odt": 23, "xml": 19,
}
BUILTIN_STYLES = {
    "normal": -1, "heading 1": -2, "heading 2": -3, "heading 3": -4, "heading 4": -5,
    "heading 5": -6, "heading 6": -7, "heading 7": -8, "heading 8": -9, "heading 9": -10,
    "title": -63, "subtitle": -75, "list bullet": -49, "list number": -50,
    "list paragraph": -180, "quote": -181, "intense quote": -182, "no spacing": -158,
}
HIGHLIGHT = {
    "none": 0, "black": 1, "blue": 2, "turquoise": 3, "brightgreen": 4, "green": 11,
    "pink": 5, "red": 6, "yellow": 7, "white": 8, "darkblue": 9, "teal": 10,
    "violet": 12, "darkred": 13, "darkyellow": 14, "gray50": 15, "gray25": 16,
}
ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
COMMENT_LIMIT = 500
CM_TO_PT = 72.0 / 2.54


class WordError(Exception):
    pass


def com_message(e: pywintypes.com_error) -> str:
    hr, msg, exc, _ = e.args
    if exc and len(exc) > 2 and exc[2]:
        return f"Word: {exc[2]} (0x{hr & 0xFFFFFFFF:08X})"
    return f"COM error 0x{hr & 0xFFFFFFFF:08X}: {msg}"


class ComWorker:
    """One STA thread owns every Word COM object."""

    def __init__(self) -> None:
        self._jobs: queue.Queue = queue.Queue()
        self._app = None
        threading.Thread(target=self._loop, name="word-com", daemon=True).start()

    def _loop(self) -> None:
        pythoncom.CoInitialize()
        while True:
            fn, fut, edit = self._jobs.get()
            if not fut.set_running_or_notify_cancel():
                continue
            try:
                fut.set_result(self._call(fn))
            except BaseException as e:  # noqa: BLE001 - surfaced to the client
                fut.set_exception(e)
            finally:
                if edit:
                    after_agent_edit(self)

    def app(self):
        if self._app is None:
            try:
                self._app = win32com.client.dynamic.Dispatch(
                    win32com.client.GetActiveObject("Word.Application")._oleobj_)
            except pywintypes.com_error as e:
                if e.args[0] != MK_E_UNAVAILABLE:
                    raise
                self._app = win32com.client.dynamic.Dispatch("Word.Application")
                self._app.Visible = True
        return self._app

    def _call(self, fn: Callable[[Any], Any]):
        deadline = time.monotonic() + 5.0
        reconnected = False
        while True:
            try:
                return fn(self.app())
            except pywintypes.com_error as e:
                hr = e.args[0]
                if hr in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER) and time.monotonic() < deadline:
                    time.sleep(0.05)  # Word busy (dialog open or mid-edit)
                    continue
                if hr in (RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED) and not reconnected:
                    self._app, reconnected = None, True  # Word was restarted
                    continue
                if hr in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
                    raise WordError("Word is busy: close any open dialog in Word and retry") from None
                raise WordError(com_message(e)) from None

    async def run(self, fn: Callable[[Any], Any], edit: bool = False):
        """Run fn(app) on the COM thread. edit=True marks an agent edit: afterwards the
        change baselines are refreshed so word_changes only reports other people's edits."""
        fut: concurrent.futures.Future = concurrent.futures.Future()
        self._jobs.put((fn, fut, edit))
        try:
            return await asyncio.wait_for(asyncio.wrap_future(fut), CALL_TIMEOUT)
        except asyncio.TimeoutError:
            raise WordError(f"Word did not answer within {CALL_TIMEOUT:.0f}s (a dialog may be open in Word)") from None


worker = ComWorker()
mcp = FastMCP("word", log_level="WARNING", instructions=(
    "Works live inside the Microsoft Word open on this PC (local files and SharePoint/OneDrive documents "
    "opened in Word desktop); edits appear instantly in the user's window.\n"
    "Workflow: 1) word_context reads the whole live document (text, structure, images, comments, the user's "
    "unsaved edits) with Word paragraph numbers [pN]. 2) Write with word_insert_markdown - a whole section "
    "(headings, lists, code blocks, tables, images with captions) in ONE call; rewrite a section with "
    "word_replace_section. 3) Target places by `anchor` (a heading's or paragraph's text) rather than numbers "
    "when the user may be editing too; pass `expect` with numbers. 4) Before further edits, word_changes "
    "shows what the user changed since your last read (cheap). 5) Check layout with word_page_image.\n"
    "NEVER save (word_save / word_close save=true) unless the user explicitly asks: saving is the user's "
    "decision (they press Ctrl+S). Every edit is one Ctrl+Z step. Several small steps: word_batch.\n"
    "Anything without a tool: word_api_search + word_com (full Word object model). Most tools take an "
    "optional `document` (name or full path/URL); default is the active document."))


# --- helpers (run on the COM thread) -------------------------------------------

def get_doc(app, document: str | None):
    if app.Documents.Count == 0:
        raise WordError("No document is open in Word (use word_new_document or word_open)")
    if not document:
        return app.ActiveDocument
    want = document.strip().lower()
    for i in range(1, app.Documents.Count + 1):
        d = app.Documents(i)
        if want in (d.Name.lower(), d.FullName.lower()):
            return d
    raise WordError(f"No open document named {document!r}")


def clean(text: str) -> str:
    """Word story text -> plain text (cell/row markers, field marks, line breaks)."""
    text = text.replace("\r\x07", "\t").replace("\x07", "")
    text = re.sub("[\x13\x14\x15]", "", text)
    text = text.replace("\x0b", "\n").replace("\x0c", "\n").replace("\x1e", "-").replace("\x1f", "")
    return text.replace("\r", "\n")


def to_word(text: str) -> str:
    return text.replace("\r\n", "\r").replace("\n", "\r")


def para_index(doc, pos: int) -> int:
    # +1 so a position at a paragraph's first character counts that paragraph
    return doc.Range(0, pos + 1).Paragraphs.Count


def check_para(doc, n: int) -> int:
    count = doc.Paragraphs.Count
    if not 1 <= n <= count:
        raise WordError(f"paragraph {n} is out of range (document has {count})")
    return count


def plain_paragraphs(doc, a: int, b: int):
    """New paragraphs start as Normal without list numbering, whatever paragraph they were split from."""
    r = doc.Range(a, b)
    r.Style = doc.Styles(BUILTIN_STYLES["normal"])
    r.ListFormat.RemoveNumbers()


def resolve_style(doc, style: str):
    key = style.strip().lower()
    try:
        return doc.Styles(style)
    except pywintypes.com_error:
        if key in BUILTIN_STYLES:  # localized Word: fall back to the built-in id
            return doc.Styles(BUILTIN_STYLES[key])
        raise WordError(f"style {style!r} not found")


def color_bgr(value: str) -> int:
    v = value.strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", v):
        raise WordError(f"color must be #RRGGBB, got {value!r}")
    r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    return r | (g << 8) | (b << 16)


def find_ranges(doc, text: str, match_case: bool, whole_word: bool, limit: int):
    """All matches of `text` in the main story, in order."""
    out = []
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    while len(out) < limit:
        # FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        # MatchAllWordForms, Forward, Wrap(wdFindStop=0)
        if not f.Execute(text, match_case, whole_word, False, False, False, True, 0):
            break
        out.append((rng.Start, rng.End))
        rng.Collapse(WD_COLLAPSE_END)
    return out


def targets(doc, app, paragraph, end_paragraph, find, occurrence, match_case, start, end, selection):
    """Resolve the target spec shared by the formatting/comment tools to (start, end) spans."""
    if selection:
        s = app.Selection
        return [(s.Start, s.End)]
    if start is not None and end is not None:
        return [(int(start), int(end))]
    if paragraph is not None:
        check_para(doc, paragraph)
        last = end_paragraph or paragraph
        check_para(doc, last)
        a = doc.Paragraphs(paragraph).Range.Start
        b = doc.Paragraphs(last).Range.End
        return [(a, b)]
    if find:
        spans = find_ranges(doc, find, match_case, False, 10000)
        if not spans:
            raise WordError(f"text {find!r} not found")
        if occurrence:
            if occurrence > len(spans):
                raise WordError(f"only {len(spans)} matches of {find!r}")
            return [spans[occurrence - 1]]
        return spans
    raise WordError("give a target: paragraph, find, start+end, or selection=true")


class undo_step:
    """Group one tool's edits into a single Ctrl+Z step. freeze=True also pauses screen updates,
    which only pays off for bulk edits (the repaint on resume costs ~10 ms or more)."""

    def __init__(self, app, label: str, freeze: bool = False):
        self.app, self.label, self.freeze = app, label, freeze

    def __enter__(self):
        try:
            self.app.UndoRecord.StartCustomRecord(f"MCP: {self.label}")
        except pywintypes.com_error:
            pass
        if self.freeze:
            self.app.ScreenUpdating = False
        return self

    def __exit__(self, *exc):
        if self.freeze:
            self.app.ScreenUpdating = True
        try:
            self.app.UndoRecord.EndCustomRecord()
        except pywintypes.com_error:
            pass
        return False


def count_text(doc, text: str, match_case: bool, whole_word: bool) -> int:
    """Match count from one read of the story text (instead of one Find round trip per match)."""
    pat = re.escape(text)
    if whole_word:
        pat = r"(?<!\w)" + pat + r"(?!\w)"
    return len(re.findall(pat, doc.Content.Text, 0 if match_case else re.I))


def py_value(v: Any) -> Any:
    if isinstance(v, (str, int, float, bool)) or v is None:
        return v
    if isinstance(v, (tuple, list)):
        return [py_value(x) for x in v]
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if hasattr(v, "_oleobj_"):
        try:
            ti = v._oleobj_.GetTypeInfo()
            return {"com_object": ti.GetDocumentation(-1)[0]}
        except pywintypes.com_error:
            return {"com_object": "IDispatch"}
    return str(v)


# --- change baselines, context cache, anchors (COM thread) ------------------------

SNAPSHOTS: dict[str, dict] = {}   # FullName -> what the agent last saw (word_changes baseline)
CTX_CACHE: dict[str, tuple] = {}  # FullName -> (fingerprint, Context)
STYLES_READY: set[str] = set()    # documents whose built-in styles were instantiated
EDITS = [0]                       # bumps on every agent edit (part of the context fingerprint)


def paragraph_texts(d) -> list[str]:
    """Every paragraph's plain text from one Content.Text read (table cells count, as in Word)."""
    parts = d.Content.Text.split("\r")
    if parts and parts[-1] == "":
        parts.pop()
    return [clean(p).strip() for p in parts]


def doc_counts(d) -> dict:
    return {"images": d.InlineShapes.Count + d.Shapes.Count, "comments": d.Comments.Count,
            "tables": d.Tables.Count, "tracked_changes": d.Revisions.Count}


def take_snapshot(d, texts: list[str] | None = None) -> None:
    SNAPSHOTS[d.FullName] = {"texts": texts if texts is not None else paragraph_texts(d), "counts": doc_counts(d)}


def after_agent_edit(w: "ComWorker") -> None:
    """The agent's own edits must not show up in word_changes: re-baseline every tracked document."""
    EDITS[0] += 1
    if not SNAPSHOTS or w._app is None:
        return
    try:
        app = w._app
        open_docs = {app.Documents(i).FullName: app.Documents(i) for i in range(1, app.Documents.Count + 1)}
        for name in list(SNAPSHOTS):
            if name in open_docs:
                take_snapshot(open_docs[name])
            else:
                SNAPSHOTS.pop(name)
    except pywintypes.com_error:
        pass


def get_context(d, refresh: bool = False) -> Context:
    """Parsed WordOpenXML, reused while the document is unchanged (text, images, comments,
    tracked changes and agent edits are all part of the fingerprint)."""
    fp = (hash(d.Content.Text), d.InlineShapes.Count, d.Shapes.Count, d.Comments.Count,
          d.Revisions.Count, EDITS[0])
    hit = CTX_CACHE.get(d.FullName)
    if hit and hit[0] == fp and not refresh:
        return hit[1]
    ctx = Context(d.WordOpenXML)
    CTX_CACHE[d.FullName] = (fp, ctx)
    return ctx


def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t.strip().lstrip("#").strip()).lower()


def locate(d, paragraph: int | None = None, anchor: str = "", expect: str = "",
           texts: list[str] | None = None) -> int:
    """Paragraph number from a number or an anchor text (a heading or any paragraph), checking
    `expect` so an edit never lands on a paragraph that changed since the agent read it."""
    texts = texts if texts is not None else paragraph_texts(d)
    if anchor:
        want = _norm(anchor)
        hits = [i + 1 for i, t in enumerate(texts) if _norm(t) == want]
        hits = hits or [i + 1 for i, t in enumerate(texts) if _norm(t).startswith(want)]
        hits = hits or [i + 1 for i, t in enumerate(texts) if want in _norm(t)]
        if not hits:
            raise WordError(f"no paragraph matches anchor {anchor!r}; re-read with word_context")
        n = hits[0]
    elif paragraph:
        n = paragraph
        if not 1 <= n <= len(texts):
            raise WordError(f"paragraph {n} is out of range (document has {len(texts)})")
    else:
        raise WordError("give paragraph or anchor")
    if expect and not _norm(texts[n - 1]).startswith(_norm(expect)):
        raise WordError(f"paragraph {n} now reads {texts[n - 1][:80]!r}, not {expect!r}: the document changed "
                        "since you read it. Re-read (word_changes / word_context) or target it by anchor.")
    return n


def section_end(d, n: int) -> int | None:
    """Paragraph number where the section started by heading n ends (next heading of the same or a
    higher level), or None when it runs to the end of the document."""
    levels = get_context(d).heading_level
    if n not in levels:
        raise WordError(f"paragraph {n} is not a heading; section_end needs a heading anchor")
    lvl = levels[n]
    nxt = [m for m, l in levels.items() if m > n and l <= lvl]
    return min(nxt) if nxt else None


def new_paragraph_at(d, app, position: str, paragraph: int = 0, anchor: str = "", expect: str = "") -> int:
    """Create one empty Normal paragraph at the requested place and return its start position.
    position: end | start | after | before | section_end (anchor = heading) | cursor (after the
    paragraph holding the cursor)."""
    if position == "end":
        pos, _ = _append(d, "")
    elif position == "start":
        d.Range(0, 0).InsertBefore("\r")
        pos = 0
    elif position == "cursor":
        p = para_index(d, app.Selection.Start)
        return new_paragraph_at(d, app, "after", p)
    elif position in ("after", "before", "section_end"):
        texts = paragraph_texts(d)
        n = locate(d, paragraph or None, anchor, expect, texts)
        if position == "section_end":
            end = section_end(d, n)
            if end is None:
                return new_paragraph_at(d, app, "end")
            return new_paragraph_at(d, app, "before", end)
        if position == "after" and n == len(texts):
            pos, _ = _append(d, "")
        else:
            pos = d.Paragraphs(n + 1 if position == "after" else n).Range.Start
            d.Range(pos, pos).InsertBefore("\r")
    else:
        raise WordError("position must be end, start, after, before, section_end or cursor")
    plain_paragraphs(d, pos, pos)
    return pos


def prepare_styles(d) -> None:
    """Instantiate the built-in styles a markdown fragment uses, so the inserted XML takes the
    document's own (template) look instead of the fragment's fallback definitions."""
    if d.FullName in STYLES_READY:
        return
    for sid in (-2, -3, -4, -5, -6, -7, -63, -181, -35, -180, -86):
        try:
            s = d.Styles(sid)
            s.Priority = s.Priority
        except pywintypes.com_error:
            pass
    try:
        s = d.Styles("Table Grid")
        s.Priority = s.Priority
    except pywintypes.com_error:
        pass
    STYLES_READY.add(d.FullName)


def usable_width_cm(d, pos: int) -> float:
    ps = d.Range(pos, pos).Sections(1).PageSetup
    return max(4.0, (ps.PageWidth - ps.LeftMargin - ps.RightMargin) / CM_TO_PT)


def autosave_state(d):
    try:
        return bool(d.AutoSaveOn)
    except (pywintypes.com_error, AttributeError):
        return None


# --- tools ---------------------------------------------------------------------

@mcp.tool()
async def word_status() -> dict:
    """Word version, open documents, the active one with its counts, and the cursor position."""
    def fn(app):
        docs = []
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents(i)
            docs.append({"name": d.Name, "path": d.FullName, "saved": bool(d.Saved),
                         "cloud": d.FullName.lower().startswith("http"), "autosave": autosave_state(d)})
        out = {"version": app.Version, "build": app.Build, "documents": docs}
        if docs:
            d = app.ActiveDocument
            s = app.Selection
            out["active"] = {
                "name": d.Name, "path": d.FullName, "paragraphs": d.Paragraphs.Count,
                "characters": d.Characters.Count, "tables": d.Tables.Count,
                "comments": d.Comments.Count, "revisions": d.Revisions.Count,
                "track_changes": bool(d.TrackRevisions),
                "selection": {"start": s.Start, "end": s.End, "text": clean(s.Text)[:200]},
            }
        return out
    return await worker.run(fn)


@mcp.tool()
async def word_new_document(text: str = "", template: str = "") -> dict:
    """Create a new document (optionally from a template path) and optionally fill it with text (newlines = paragraphs)."""
    def fn(app):
        d = app.Documents.Add(template) if template else app.Documents.Add()
        if text:
            d.Content.Text = to_word(text)
        return {"name": d.Name, "paragraphs": d.Paragraphs.Count}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_open(path: str, read_only: bool = False) -> dict:
    """Open a document (.docx, .doc, .rtf, .txt, .pdf, ...) from disk or a SharePoint / OneDrive URL
    (https://...), and make it active. If it is already open it is just activated.
    Note: like Word's own File > Open, this replaces a lone, untouched blank new document
    (reported as replaced_blank_document)."""
    def fn(app):
        url = re.match(r"https?://", path.strip(), re.I) is not None
        full = path.strip() if url else os.path.abspath(path)
        if not url and not os.path.exists(full):
            raise WordError(f"file not found: {full}")
        for i in range(1, app.Documents.Count + 1):
            if app.Documents(i).FullName.lower() == full.lower():
                d = app.Documents(i)
                d.Activate()
                return {"name": d.Name, "path": d.FullName, "paragraphs": d.Paragraphs.Count, "already_open": True}
        blank = None
        if app.Documents.Count == 1:
            only = app.Documents(1)
            if not only.Path and only.Saved and only.Content.End <= 1:
                blank = only.Name
        prev = app.DisplayAlerts
        app.DisplayAlerts = WD_ALERTS_NONE
        try:
            d = app.Documents.Open(full, False, read_only, True)
        finally:
            app.DisplayAlerts = prev
        d.Activate()
        out = {"name": d.Name, "path": d.FullName, "paragraphs": d.Paragraphs.Count}
        if blank and all(app.Documents(i).Name != blank for i in range(1, app.Documents.Count + 1)):
            out["replaced_blank_document"] = blank
        return out
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_save(path: str = "", format: str = "", document: str = "") -> dict:
    """Save the document. With `path`, save as that file; `format` (docx, pdf, rtf, txt, html, odt, ...)
    defaults to the path's extension. PDF/XPS export leaves the open document unchanged."""
    def fn(app):
        d = get_doc(app, document)
        if not path:
            if not d.Path:
                raise WordError("document was never saved: give a path")
            d.Save()
            return {"saved": d.FullName}
        full = os.path.abspath(path)
        fmt = (format or os.path.splitext(full)[1].lstrip(".")).lower()
        if fmt not in WD_FORMAT:
            raise WordError(f"unknown format {fmt!r}; use one of {', '.join(sorted(WD_FORMAT))}")
        prev = app.DisplayAlerts
        app.DisplayAlerts = WD_ALERTS_NONE
        try:
            if fmt in ("pdf", "xps"):
                d.ExportAsFixedFormat(full, WD_FORMAT[fmt])
            else:
                d.SaveAs2(full, WD_FORMAT[fmt])
        finally:
            app.DisplayAlerts = prev
        return {"saved": full, "format": fmt, "size_bytes": os.path.getsize(full)}
    return await worker.run(fn)


@mcp.tool()
async def word_close(save: bool = False, document: str = "") -> dict:
    """Close a document. save=false discards unsaved changes, save=true saves first (needs a path)."""
    def fn(app):
        d = get_doc(app, document)
        name = d.Name
        if save and not d.Path:
            raise WordError("document was never saved: call word_save with a path first")
        d.Close(-1 if save else 0)
        return {"closed": name, "saved": save}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_activate(document: str) -> dict:
    """Bring an open document to the front and make it the active one."""
    def fn(app):
        d = get_doc(app, document)
        d.Activate()
        return {"active": d.Name}
    return await worker.run(fn)


@mcp.tool()
async def word_read(start_paragraph: int = 1, end_paragraph: int = 0, with_styles: bool = False,
                    document: str = "") -> dict:
    """Read paragraphs as [{n, text}] (1-based). end_paragraph=0 reads to the end.
    with_styles adds each paragraph's style name (one extra Word call per paragraph)."""
    def fn(app):
        d = get_doc(app, document)
        count = d.Paragraphs.Count
        last = count if end_paragraph <= 0 else min(end_paragraph, count)
        first = max(1, start_paragraph)
        if first > last:
            return {"paragraph_count": count, "paragraphs": []}
        if first == 1 and last == count:
            raw = d.Content.Text
            take_snapshot(d)  # a full read is a word_changes baseline too
        else:
            raw = d.Range(d.Paragraphs(first).Range.Start, d.Paragraphs(last).Range.End).Text
        parts = raw.split("\r")
        if parts and parts[-1] == "":
            parts.pop()
        paras = []
        for i, t in enumerate(parts):
            item = {"n": first + i, "text": clean(t)}
            paras.append(item)
        if with_styles:
            for item in paras:
                if item["n"] <= count:
                    item["style"] = d.Paragraphs(item["n"]).Style.NameLocal
        return {"paragraph_count": count, "paragraphs": paras}
    return await worker.run(fn)


@mcp.tool()
async def word_get_text(document: str = "") -> str:
    """The whole document as plain text, one paragraph per line (fastest way to read everything)."""
    def fn(app):
        return clean(get_doc(app, document).Content.Text).replace("\r", "\n").rstrip("\n")
    return await worker.run(fn)


@mcp.tool()
async def word_insert(text: str, position: str = "end", paragraph: int = 0, anchor: str = "", expect: str = "",
                      style: str = "", document: str = "") -> dict:
    """Insert plain text; newlines make new paragraphs. For formatted content (headings, lists, code,
    tables, images) use word_insert_markdown instead: one call for a whole section.
    position: "end" | "start" | "after" / "before" (paragraph N or `anchor`) | "section_end" (end of the
    section under heading `anchor`) | "cursor" (typed at the selection, inline).
    anchor: text of the target paragraph (e.g. a heading) - robust when the user edits the document too.
    expect: start of paragraph N's current text; the edit is refused if it no longer matches.
    style applies a paragraph style ("Heading 1", "List Bullet", ...); without it new paragraphs are Normal."""
    def fn(app):
        d = get_doc(app, document)
        body = to_word(text)
        with undo_step(app, "insert", freeze=len(body) > 2000):
            if position == "cursor":
                r = app.Selection.Range
                r.Text = body
                a, b = r.Start, r.End
                app.Selection.SetRange(b, b)
            else:
                a = new_paragraph_at(d, app, position, paragraph, anchor, expect)
                d.Range(a, a).InsertAfter(body)
                b = a + len(body)
                if style:
                    d.Range(a, b).Style = resolve_style(d, style)
        return {"first_paragraph": para_index(d, a), "last_paragraph": para_index(d, b),
                "paragraph_count": d.Paragraphs.Count}
    return await worker.run(fn, edit=True)


def _append(d, body: str) -> tuple[int, int]:
    last = d.Paragraphs.Last.Range
    if last.End - last.Start > 1:  # last paragraph has text: start a new one
        d.Content.InsertParagraphAfter()
        last = d.Paragraphs.Last.Range
    pos = last.Start
    d.Range(pos, pos).InsertAfter(body)
    return pos, pos + len(body)


@mcp.tool()
async def word_set_paragraph(text: str, paragraph: int = 0, anchor: str = "", expect: str = "",
                             document: str = "") -> dict:
    """Replace one paragraph's text (by number or `anchor` text), keeping its style and formatting.
    expect: start of the paragraph's current text; refused if the document changed."""
    def fn(app):
        d = get_doc(app, document)
        n = locate(d, paragraph or None, anchor, expect)
        r = d.Paragraphs(n).Range
        with undo_step(app, "set paragraph"):
            d.Range(r.Start, r.End - 1).Text = to_word(text)
        return {"paragraph": n, "paragraph_count": d.Paragraphs.Count}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_delete_paragraphs(start_paragraph: int = 0, end_paragraph: int = 0, anchor: str = "",
                                 expect: str = "", document: str = "") -> dict:
    """Delete paragraphs start..end (inclusive; end defaults to start). `anchor` picks the start paragraph
    by its text; `expect` refuses the delete if the start paragraph no longer reads like that."""
    def fn(app):
        d = get_doc(app, document)
        first = locate(d, start_paragraph or None, anchor, expect)
        last = max(first, end_paragraph or first)
        check_para(d, last)
        with undo_step(app, "delete paragraphs", freeze=last - first > 20):
            d.Range(d.Paragraphs(first).Range.Start, d.Paragraphs(last).Range.End).Delete()
        return {"deleted": last - first + 1, "paragraph_count": d.Paragraphs.Count}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_find(text: str, match_case: bool = False, whole_word: bool = False, max_results: int = 50,
                    document: str = "") -> dict:
    """Find text; returns each match's paragraph number, character span and surrounding context."""
    def fn(app):
        d = get_doc(app, document)
        spans = find_ranges(d, text, match_case, whole_word, max_results)
        out = []
        for a, b in spans:
            ctx = d.Range(max(0, a - 40), b + 40).Text
            out.append({"paragraph": para_index(d, a), "start": a, "end": b, "context": clean(ctx).replace("\n", " / ")})
        return {"count": len(out), "matches": out}
    return await worker.run(fn)


@mcp.tool()
async def word_replace(find: str, replace: str, replace_all: bool = True, match_case: bool = False,
                       whole_word: bool = False, use_wildcards: bool = False, document: str = "") -> dict:
    """Find and replace in the document body (Word's own Replace, so it keeps formatting).
    Returns how many replacements were made."""
    def fn(app):
        d = get_doc(app, document)
        before = None if use_wildcards or "^" in find else count_text(d, find, match_case, whole_word)
        with undo_step(app, "replace", freeze=replace_all):
            r = d.Content
            f = r.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            # ..., Forward, Wrap(wdFindContinue=1), Format, ReplaceWith, Replace(wdReplaceAll=2 / One=1)
            done = f.Execute(find, match_case, whole_word, use_wildcards, False, False, True, 1, False,
                             to_word(replace), 2 if replace_all else 1)
        count = (before if replace_all else min(1, before)) if before is not None else None
        return {"replaced": bool(done), "count": count if done else 0}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_format(paragraph: int | None = None, end_paragraph: int | None = None, find: str = "",
                      occurrence: int = 0, match_case: bool = False, start: int | None = None,
                      end: int | None = None, selection: bool = False,
                      bold: bool | None = None, italic: bool | None = None, underline: bool | None = None,
                      strikethrough: bool | None = None, font: str = "", size: float | None = None,
                      color: str = "", highlight: str = "", style: str = "", align: str = "",
                      space_after: float | None = None, line_spacing: float | None = None,
                      anchor: str = "", expect: str = "", document: str = "") -> dict:
    """Format text. Target: paragraph (to end_paragraph) or `anchor` (a paragraph's text), or every match
    of `find` (or just `occurrence` N), or start+end character positions, or selection=true.
    expect: start of the target paragraph's current text; refused if the document changed.
    Character: bold, italic, underline, strikethrough, font, size (pt), color (#RRGGBB),
    highlight (yellow, green, brightgreen, turquoise, pink, red, blue, gray25, none...).
    Paragraph: style, align (left/center/right/justify), space_after (pt), line_spacing (pt)."""
    def fn(app):
        nonlocal paragraph
        d = get_doc(app, document)
        if anchor or (expect and paragraph):
            paragraph = locate(d, paragraph, anchor, expect)
        if highlight and highlight.lower() not in HIGHLIGHT:
            raise WordError(f"highlight must be one of {', '.join(HIGHLIGHT)}")
        if align and align.lower() not in ALIGN:
            raise WordError("align must be left, center, right or justify")
        st = resolve_style(d, style) if style else None
        rgb = color_bgr(color) if color else None

        def apply_font(f):
            if bold is not None: f.Bold = bold
            if italic is not None: f.Italic = italic
            if underline is not None: f.Underline = 1 if underline else 0
            if strikethrough is not None: f.StrikeThrough = strikethrough
            if font: f.Name = font
            if size is not None: f.Size = size
            if rgb is not None: f.Color = rgb

        def apply_para(pf):
            if align: pf.Alignment = ALIGN[align.lower()]
            if space_after is not None: pf.SpaceAfter = space_after
            if line_spacing is not None:
                pf.LineSpacingRule = 4  # wdLineSpaceExactly
                pf.LineSpacing = line_spacing

        para = align or space_after is not None or line_spacing is not None
        if find and not occurrence and paragraph is None and start is None and not selection:
            # Every match: one Find/Replace with formatting ("^&" = the found text), ~30x faster
            # than formatting match by match.
            n = count_text(d, find, match_case, False) if "^" not in find else None
            if n == 0:
                raise WordError(f"text {find!r} not found")
            with undo_step(app, "format", freeze=True):
                r = d.Content
                fnd = r.Find
                fnd.ClearFormatting()
                rep = fnd.Replacement
                rep.ClearFormatting()
                if st is not None: rep.Style = st
                apply_font(rep.Font)
                if para: apply_para(rep.ParagraphFormat)
                prev_hl = None
                if highlight:
                    prev_hl = app.Options.DefaultHighlightColorIndex
                    app.Options.DefaultHighlightColorIndex = HIGHLIGHT[highlight.lower()]
                    rep.Highlight = highlight.lower() != "none"
                try:
                    fnd.Execute(find, match_case, False, False, False, False, True, 0, True, "^&", 2)
                finally:
                    if prev_hl is not None:
                        app.Options.DefaultHighlightColorIndex = prev_hl
                    rep.ClearFormatting()
            return {"formatted_matches": n}

        spans = targets(d, app, paragraph, end_paragraph, find, occurrence, match_case, start, end, selection)
        with undo_step(app, "format", freeze=len(spans) > 20):
            for a, b in spans:
                r = d.Range(a, b)
                if st is not None: r.Style = st
                apply_font(r.Font)
                if highlight: r.HighlightColorIndex = HIGHLIGHT[highlight.lower()]
                if para: apply_para(r.ParagraphFormat)
        return {"formatted_ranges": len(spans)}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_insert_table(rows: list[list[str]], after_paragraph: int = 0, header: bool = True,
                            style: str = "Table Grid", autofit: bool = True, document: str = "") -> dict:
    """Insert a table from a 2D list of cell strings (first row = header when header=true).
    Placed after paragraph N, or at the end when after_paragraph=0."""
    def fn(app):
        if not rows or not rows[0]:
            raise WordError("rows must be a non-empty 2D list")
        d = get_doc(app, document)
        cols = max(len(r) for r in rows)
        cell = lambda v: str(v).replace("\t", " ").replace("\r", " ").replace("\n", " ")
        body = "\r".join("\t".join(cell(v) for v in (r + [""] * (cols - len(r)))) for r in rows)
        with undo_step(app, "insert table", freeze=True):
            if after_paragraph and after_paragraph < d.Paragraphs.Count:
                check_para(d, after_paragraph)
                pos = d.Paragraphs(after_paragraph + 1).Range.Start
                d.Range(pos, pos).InsertBefore(body + "\r")
                a, b = pos, pos + len(body)
            else:
                a, b = _append(d, body)
            # One call builds the whole table instead of one call per cell.
            plain_paragraphs(d, a, b)  # a table after a list must not inherit its numbering
            t = d.Range(a, b).ConvertToTable(WD_SEPARATE_BY_TABS, len(rows), cols)
            if style:
                try:
                    t.Style = style
                except pywintypes.com_error:
                    pass
            if header:
                t.Rows(1).HeadingFormat = True
                t.Rows(1).Range.Font.Bold = True
            if autofit:
                t.AutoFitBehavior(1)  # wdAutoFitContent
        idx = 0
        for i in range(1, d.Tables.Count + 1):
            if d.Tables(i).Range.Start == t.Range.Start:
                idx = i
                break
        return {"table": idx, "rows": len(rows), "columns": cols}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_read_tables(table: int = 0, document: str = "") -> dict:
    """Read tables as 2D lists of cell text. table=0 reads all, N reads table N (1-based)."""
    def fn(app):
        d = get_doc(app, document)
        n = d.Tables.Count
        which = [table] if table else list(range(1, n + 1))
        out = []
        for i in which:
            if not 1 <= i <= n:
                raise WordError(f"table {i} out of range (document has {n})")
            t = d.Tables(i)
            nr, nc = t.Rows.Count, t.Columns.Count
            grid = None
            if t.Uniform:
                cells = t.Range.Text.split("\r\x07")
                # every row is nc cells plus an empty end-of-row marker
                if len(cells) >= nr * (nc + 1):
                    grid = [[clean(cells[r * (nc + 1) + c]) for c in range(nc)] for r in range(nr)]
            if grid is None:  # merged cells: walk them (slower, exact)
                grid = [[""] * nc for _ in range(nr)]
                cs = t.Range.Cells
                for k in range(1, cs.Count + 1):
                    c = cs(k)
                    grid[c.RowIndex - 1][c.ColumnIndex - 1] = clean(c.Range.Text[:-2])
            out.append({"table": i, "rows": nr, "columns": nc, "first_paragraph": para_index(d, t.Range.Start),
                        "cells": grid})
        return {"table_count": n, "tables": out}
    return await worker.run(fn)


@mcp.tool()
async def word_set_cell(table: int, row: int, column: int, text: str, document: str = "") -> dict:
    """Set one table cell's text (1-based table/row/column)."""
    def fn(app):
        d = get_doc(app, document)
        with undo_step(app, "set cell"):
            d.Tables(table).Cell(row, column).Range.Text = text
        return {"table": table, "row": row, "column": column}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_add_comment(comment: str, paragraph: int | None = None, find: str = "", occurrence: int = 1,
                           selection: bool = False, document: str = "") -> dict:
    """Add a comment on a paragraph, on the Nth match of `find` (default first), or on the selection.
    Returns the comment's number n (as used by word_reply_comment / word_resolve_comment)."""
    def fn(app):
        d = get_doc(app, document)
        spans = targets(d, app, paragraph, None, find, occurrence if find else 0, False, None, None, selection)
        a, b = spans[0]
        if paragraph is not None:
            b -= 1  # leave the paragraph mark out of the commented text
        with undo_step(app, "comment"):
            c = d.Comments.Add(d.Range(a, b), comment)
        return {"n": c.Index, "comment_count": d.Comments.Count, "paragraph": para_index(d, a)}
    return await worker.run(fn, edit=True)


def get_comment(d, n: int):
    count = d.Comments.Count
    if not 1 <= n <= count:
        raise WordError(f"comment {n} does not exist (document has {count}); see word_list_comments")
    return d.Comments(n)


def comment_brief(c) -> dict:
    return {"n": c.Index, "author": c.Author, "text": clean(c.Range.Text).strip()}


@mcp.tool()
async def word_list_comments(include_resolved: bool = True, document: str = "") -> dict:
    """All comment threads: each top-level comment with author, date, text, the text it is on, its paragraph,
    resolved state and its replies. `n` numbers are what the other comment tools take."""
    def fn(app):
        d = get_doc(app, document)
        ctx = get_context(d)  # one WordOpenXML read covers every comment, reply and resolved flag
        flat = ctx.comments()
        if len(flat) != d.Comments.Count:
            raise WordError("comment numbering mismatch between Word and its XML; retry")
        threads, by_n = [], {}
        for c in flat:
            item = {k: c[k] for k in ("n", "author", "date", "text")}
            if c["reply_to"] and c["reply_to"] in by_n:
                by_n[c["reply_to"]]["replies"].append(item)
                continue
            item.update({"on": c["on"], "paragraph": c["paragraph"], "resolved": c["resolved"], "replies": []})
            by_n[c["n"]] = item
            threads.append(item)
        shown = [t for t in threads if include_resolved or not t["resolved"]]
        return {"comment_count": len(flat), "threads": len(threads),
                "open_threads": sum(not t["resolved"] for t in threads), "comments": shown}
    return await worker.run(fn)


@mcp.tool()
async def word_reply_comment(n: int, text: str, document: str = "") -> dict:
    """Reply to comment n (a reply to a reply goes on the same thread). Returns the new reply's n."""
    def fn(app):
        d = get_doc(app, document)
        c = get_comment(d, n)
        anc = c.Ancestor
        root = anc if anc is not None else c
        with undo_step(app, "reply"):
            r = root.Replies.Add(root.Range, text)
        return {"n": r.Index, "reply_to": comment_brief(root), "thread_replies": root.Replies.Count}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_resolve_comment(n: int, resolved: bool = True, document: str = "") -> dict:
    """Mark comment n's thread resolved (or reopen it with resolved=false)."""
    def fn(app):
        d = get_doc(app, document)
        c = get_comment(d, n)
        anc = c.Ancestor
        root = anc if anc is not None else c
        with undo_step(app, "resolve comment"):
            root.Done = resolved
        return {"thread": comment_brief(root), "resolved": bool(root.Done)}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_edit_comment(n: int, text: str, document: str = "") -> dict:
    """Replace the text of comment (or reply) n."""
    def fn(app):
        d = get_doc(app, document)
        c = get_comment(d, n)
        with undo_step(app, "edit comment"):
            c.Range.Text = text
        return comment_brief(c)
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_delete_comment(n: int = 0, document: str = "") -> dict:
    """Delete comment n with its replies, or every comment when n=0."""
    def fn(app):
        d = get_doc(app, document)
        with undo_step(app, "delete comment"):
            if n:
                get_comment(d, n).Delete()
            else:
                # Document.DeleteAllComments fails over COM ("not available"); delete from the end.
                for i in range(d.Comments.Count, 0, -1):
                    d.Comments(i).Delete()
        return {"comment_count": d.Comments.Count}
    return await worker.run(fn, edit=True)


# --- whole-document context and images ---------------------------------------------

IMAGE_FORMATS = {"image/png": "png", "image/jpeg": "jpeg", "image/gif": "gif", "image/webp": "webp"}


def to_mcp_image(ctype: str, data: bytes, max_px: int) -> tuple[Image, dict]:
    """Pass PNG/JPEG/GIF through when small enough; convert EMF/WMF/TIFF/BMP and shrink big ones."""
    im = PILImage.open(io.BytesIO(data))
    if ctype in ("image/x-emf", "image/x-wmf", "image/emf", "image/wmf") or im.format in ("WMF", "EMF"):
        im.load(dpi=144)
    w, h = im.size
    if ctype in IMAGE_FORMATS and max(w, h) <= max_px and len(data) < 4_000_000:
        return Image(data=data, format=IMAGE_FORMATS[ctype]), {"pixels": [w, h]}
    if max(w, h) > max_px:
        im.thumbnail((max_px, max_px))
    if im.mode not in ("RGB", "RGBA", "L"):
        im = im.convert("RGBA")
    buf = io.BytesIO()
    im.save(buf, "PNG", optimize=True)
    return Image(data=buf.getvalue(), format="png"), {"pixels": list(im.size), "converted_from": ctype}


@mcp.tool()
async def word_context(start_paragraph: int = 1, max_chars: int = 60000, include_images: bool = False,
                       max_images: int = 6, stats: bool = True, refresh: bool = False, document: str = ""):
    """READ THIS FIRST. Everything about the live document in one call (including the user's unsaved
    typing): stats, cursor, save/AutoSave state, header/footer, the full body as markdown with Word paragraph
    numbers [pN] (headings, lists, tables, bold/italic, links, tracked changes {+ins+}/{-del-}, images
    ![image n: alt]), every comment thread with replies and resolved state, and footnotes.
    Reuses its parse while the document is unchanged; refresh=true forces a re-read (e.g. after the user
    only changed formatting). stats=false skips the page/word count (faster on very long documents).
    Long documents: continue with start_paragraph from the note at the end.
    include_images=true attaches the first max_images embedded images so you can see them.
    Afterwards word_changes reports only what changed since this read."""
    def fn(app):
        d = get_doc(app, document)
        ctx = get_context(d, refresh)
        sel = app.Selection
        same_doc = app.ActiveDocument.FullName == d.FullName
        info = {
            "pages": d.ComputeStatistics(2) if stats else None,
            "words": d.ComputeStatistics(0) if stats else None,
            "paragraphs": d.Paragraphs.Count, "saved": bool(d.Saved), "autosave": autosave_state(d),
            "track_changes": bool(d.TrackRevisions),
            "cursor": (para_index(d, sel.Start), clean(sel.Text).strip()[:200]) if same_doc else None,
        }
        take_snapshot(d)
        images = []
        if include_images:
            for meta in ctx.images[:max_images]:
                try:
                    _, ctype, data = ctx.image_bytes(meta["n"])
                    images.append((meta, ctype, data))
                except IndexError:
                    pass
        return d.Name, d.FullName, bool(d.Path), ctx, info, images

    name, path, has_path, ctx, info, images = await worker.run(fn)
    body, next_para = ctx.render(start_paragraph, max_chars)
    comments = ctx.comments()
    threads = [c for c in comments if not c["reply_to"]]
    hf = ctx.headers_footers()
    props = ctx.properties()
    save_state = "saved" if info["saved"] else "UNSAVED changes (only the user decides when to save)"
    if info["autosave"]:
        save_state += ", AutoSave ON (edits upload automatically)"
    out = [f"Document: {name}  ({path if has_path else 'never saved'})  - {save_state}"]
    counts = [f"pages {info['pages']}, words {info['words']}" if info["pages"] is not None else "",
              f"paragraphs {info['paragraphs']}", f"tables {ctx.tables}", f"images {len(ctx.images)}",
              f"comment threads {len(threads)} ({sum(not c['resolved'] for c in threads)} open)",
              f"tracked changes {ctx.changes['ins']} insertions / {ctx.changes['del']} deletions"
              + (", tracking ON" if info["track_changes"] else "")]
    out.append(", ".join(c for c in counts if c))
    if info["cursor"]:
        p, s = info["cursor"]
        out.append(f"Cursor: paragraph {p}" + (f', selected "{s}"' if s else ""))
    if props:
        out.append("Properties: " + ", ".join(f"{k} {v}" for k, v in props.items()))
    for kind in ("header", "footer"):
        if hf[kind]:
            out.append(f"{kind.title()}: " + " | ".join(hf[kind]))
    if ctx.para != info["paragraphs"]:
        out.append(f"Note: paragraph numbers may drift (parsed {ctx.para}, Word reports {info['paragraphs']}); "
                   "prefer anchors, or confirm with word_read before editing by number.")
    out.append("")
    out.append("--- body ([pN] = Word paragraph number; or target paragraphs by their text with `anchor`) ---")
    out.append(body if body else "(empty)")
    if next_para:
        out.append(f"... truncated: call word_context(start_paragraph={next_para}) for the rest")
    if comments:
        out.append("")
        out.append("--- comments (n = number for word_reply_comment / word_resolve_comment) ---")
        for c in threads:
            on = f' on "{c["on"][:80]}"' if c["on"] else ""
            where = f" (p{c['paragraph']})" if c["paragraph"] else ""
            state = " [resolved]" if c["resolved"] else ""
            out.append(f"[c{c['n']}]{on}{where} {c['author']} {c['date'][:10]}{state}: {c['text']}")
            for r in comments:
                if r["reply_to"] == c["n"]:
                    out.append(f"    reply [c{r['n']}] {r['author']} {r['date'][:10]}: {r['text']}")
    notes = ctx.notes("/word/footnotes.xml", "footnote")
    if notes:
        out.append("")
        out.append("--- footnotes ---")
        out += [f"[^{i}] {t}" for i, t in notes]
    if ctx.images:
        out.append("")
        out.append("Images: " + "; ".join(f"{m['n']}: p{m['paragraph']} {m['alt_text'] or m['name']} "
                                          f"{m['width_cm']}x{m['height_cm']} cm" for m in ctx.images)
                   + " - view with word_get_image(n) or word_page_image(pages).")
    result: list = ["\n".join(out)]
    for meta, ctype, data in images:
        try:
            img, _ = to_mcp_image(ctype, data, 1024)
            result += [f"image {meta['n']} (p{meta['paragraph']}): {meta['alt_text'] or meta['name']}", img]
        except Exception as e:  # noqa: BLE001 - unreadable image format: describe instead
            result.append(f"image {meta['n']}: could not decode {ctype} ({e})")
    return result


@mcp.tool()
async def word_changes(document: str = "") -> dict:
    """What changed since you last read the document (word_context / word_read / word_changes) that you did
    not do yourself - typically the user's own edits: added, removed and changed paragraphs (with their
    current numbers) plus image / comment / table / tracked-change counts. Much cheaper than re-reading.
    The call also becomes the new baseline."""
    def fn(app):
        d = get_doc(app, document)
        now = paragraph_texts(d)
        counts = doc_counts(d)
        base = SNAPSHOTS.get(d.FullName)
        take_snapshot(d, now)
        if base is None:
            return {"baseline": "none yet - call word_context to read the document first; baseline taken now",
                    "paragraph_count": len(now)}
        import difflib
        out = []
        sm = difflib.SequenceMatcher(None, base["texts"], now, autojunk=False)
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == "equal":
                continue
            item = {"type": {"replace": "changed", "insert": "added", "delete": "removed"}[op]}
            if j2 > j1:
                item["paragraphs"] = f"{j1 + 1}" if j2 - j1 == 1 else f"{j1 + 1}-{j2}"
                item["now"] = [t[:300] for t in now[j1:j2][:20]]
            else:
                item["after_paragraph"] = j1
            if i2 > i1:
                item["before"] = [t[:300] for t in base["texts"][i1:i2][:20]]
            out.append(item)
            if len(out) >= 60:
                out.append({"type": "truncated", "note": "many changes - re-read with word_context"})
                break
        delta = {k: {"before": base["counts"][k], "now": v} for k, v in counts.items() if base["counts"].get(k) != v}
        return {"changed": bool(out or delta), "paragraph_count": len(now), "changes": out, "counts": delta,
                "note": "paragraph numbers are current; formatting-only edits are not listed"}
    return await worker.run(fn)


@mcp.tool()
async def word_insert_markdown(markdown: str, position: str = "end", paragraph: int = 0, anchor: str = "",
                               expect: str = "", base_dir: str = "", document: str = "") -> dict:
    """FASTEST way to write: a whole section (or document) of markdown in ONE call, typed straight into the
    open document with its own styles. Supports # headings, paragraphs, **bold**, *italic*, ~~strike~~,
    `inline code`, [links](https://...), - bullet / 1. numbered lists (nested; each list restarts at 1),
    > quotes, ```fenced code blocks```, | pipe | tables |, --- rules, and images:
    ![alt](C:/shots/ui.png "Caption text"){width=12cm}  - a local path (relative to base_dir), an https URL
    or a data: URI; png/jpg/gif/svg/emf. The title in quotes becomes a numbered "Figure N:" caption.
    Images wider than the page are fitted to the text width.
    position: "end" | "start" | "after" / "before" (paragraph N or `anchor` text) | "section_end"
    (end of the section under heading `anchor`) | "cursor" (after the paragraph holding the cursor).
    Nothing is saved: the user decides when to save."""
    def fn(app):
        import md_to_word
        t0 = time.perf_counter()
        d = get_doc(app, document)
        with undo_step(app, "insert markdown", freeze=True):
            prepare_styles(d)
            pos = new_paragraph_at(d, app, position, paragraph, anchor, expect)
            frag = md_to_word.build(markdown, base_dir or os.getcwd(), usable_width_cm(d, pos))
            end_before = d.Content.End
            d.Range(pos, pos).InsertXML(frag.xml)
            end = pos + (d.Content.End - end_before)
            rng = d.Range(pos, end)
            for marker, svg_path, width_cm, alt in frag.svg:  # SVGs go in natively, where XML could not
                hits = find_ranges(d, marker, True, False, 1)
                if not hits:
                    frag.warnings.append(f"SVG placeholder for {svg_path} lost")
                    continue
                r = d.Range(*hits[0])
                r.Text = ""
                shp = d.InlineShapes.AddPicture(svg_path, False, True, r)
                if shp.Width > 0:
                    ratio = shp.Height / shp.Width
                    shp.Width = width_cm * CM_TO_PT
                    shp.Height = width_cm * CM_TO_PT * ratio
                if alt:
                    shp.AlternativeText = alt
            if frag.captions:
                d.Range(pos, pos + (d.Content.End - end_before)).Fields.Update()
            end = pos + (d.Content.End - end_before)
        return {"first_paragraph": para_index(d, pos), "last_paragraph": para_index(d, max(pos, end - 1)),
                "paragraph_count": d.Paragraphs.Count, "headings": frag.headings, "tables": frag.tables,
                "images": frag.images, "captions": frag.captions, "warnings": frag.warnings,
                "ms": round((time.perf_counter() - t0) * 1000)}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_replace_section(anchor: str, markdown: str, keep_heading: bool = True, base_dir: str = "",
                               document: str = "") -> dict:
    """Rewrite one section: everything under heading `anchor` (up to the next heading of the same or a higher
    level) is replaced by `markdown`, in one call. keep_heading=false replaces the heading too (then start
    the markdown with the new heading). Sub-sections are part of the section and are replaced as well."""
    def fn(app):
        d = get_doc(app, document)
        texts = paragraph_texts(d)
        n = locate(d, None, anchor, "", texts)
        end = section_end(d, n)
        last = (end - 1) if end else len(texts)
        first = n + 1 if keep_heading else n
        removed = 0
        with undo_step(app, "replace section", freeze=True):
            if last >= first:
                # Whole paragraphs; at the document end Word keeps one empty final paragraph,
                # which the markdown then fills.
                d.Range(d.Paragraphs(first).Range.Start, d.Paragraphs(last).Range.End).Delete()
                removed = last - first + 1
        # where the new content goes: before the next section's heading, or at the end
        return {"heading_paragraph": n, "removed_paragraphs": removed,
                "insert": ("before", first) if end else ("end", 0), "doc": d.FullName}
    cut = await worker.run(fn, edit=True)
    where, para = cut.pop("insert")
    res = await word_insert_markdown(markdown, position=where, paragraph=para, base_dir=base_dir,
                                     document=cut.pop("doc"))
    return {**cut, **res}


@mcp.tool()
async def word_get_image(n: int, max_px: int = 1568, document: str = ""):
    """See embedded image n (numbering from word_context / word_list_images), at up to max_px pixels."""
    def fn(app):
        d = get_doc(app, document)
        try:
            return get_context(d).image_bytes(n)
        except IndexError as e:
            raise WordError(str(e)) from None
    meta, ctype, data = await worker.run(fn)
    img, extra = to_mcp_image(ctype, data, max_px)
    return [json.dumps({**meta, **extra, "bytes": len(data)}), img]


@mcp.tool()
async def word_list_images(document: str = "") -> dict:
    """Every embedded image and drawing: number, paragraph, alt text, size, inline or floating."""
    def fn(app):
        ctx = get_context(get_doc(app, document))
        return {"count": len(ctx.images), "images": ctx.images}
    return await worker.run(fn)


def _page_list(pages: str, count: int) -> list[int]:
    out: list[int] = []
    for part in str(pages).replace(" ", "").split(","):
        if not part:
            continue
        if part == "all":
            out += range(1, count + 1)
        elif "-" in part:
            a, b = part.split("-", 1)
            out += range(int(a or 1), min(int(b or count), count) + 1)
        else:
            out.append(int(part))
    bad = [p for p in out if not 1 <= p <= count]
    if bad:
        raise WordError(f"page {bad[0]} out of range (document has {count})")
    return out[:10]


@mcp.tool()
async def word_page_image(pages: str = "1", dpi: int = 96, document: str = ""):
    """See pages exactly as Word lays them out (text, tables, images, header/footer), as PNGs.
    pages: "1", "2-4", "1,3", "all" (max 10 per call). dpi 96 is screen size; up to 200 for small print."""
    def fn(app):
        d = get_doc(app, document)
        if d.Windows.Count == 0:
            raise WordError("document has no window")
        win = d.Windows(1)
        view = win.View.Type
        if view != 3:  # wdPrintView: pages only exist in Print Layout
            win.View.Type = 3
        try:
            pg = win.Panes(1).Pages
            count = pg.Count
            want = _page_list(pages, count)
            return count, [(p, bytes(pg(p).EnhMetaFileBits)) for p in want]
        finally:
            if view != 3:
                win.View.Type = view
    count, emfs = await worker.run(fn)
    result: list = []
    for p, emf in emfs:
        im = PILImage.open(io.BytesIO(emf))
        im.load(dpi=max(48, min(dpi, 200)))
        im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "PNG", optimize=True)
        result += [json.dumps({"page": p, "pages": count, "pixels": list(im.size)}),
                   Image(data=buf.getvalue(), format="png")]
    return result


def _image_file(path: str, data_base64: str, url: str) -> tuple[str, bool]:
    """A local file for AddPicture from a path, base64 data or an http(s) URL; (path, is_temp)."""
    import tempfile
    import urllib.request
    if path:
        full = os.path.abspath(path)
        if not os.path.isfile(full):
            raise WordError(f"file not found: {full}")
        return full, False
    if data_base64:
        head, _, b64 = data_base64.partition(",") if data_base64.startswith("data:") else ("", "", data_base64)
        data = base64.b64decode(b64)
        ext = ".svg" if b"<svg" in data[:300] else "." + (PILImage.open(io.BytesIO(data)).format or "png").lower()
    elif url:
        if not re.match(r"https?://", url, re.I):
            raise WordError("url must be http(s)")
        req = urllib.request.Request(url, headers={"User-Agent": "word-mcp"})
        with urllib.request.urlopen(req, timeout=20) as r:
            ctype = r.headers.get_content_type()
            data = r.read(25 * 2**20 + 1)
        if not ctype.startswith("image/") or len(data) > 25 * 2**20:
            raise WordError(f"not an image under 25 MB ({ctype})")
        ext = ".svg" if "svg" in ctype else "." + ctype.split("/")[1].split("+")[0]
    else:
        raise WordError("give path, data_base64 or url")
    fd, tmp = tempfile.mkstemp(suffix=ext.replace("jpeg", "jpg"))
    os.write(fd, data)
    os.close(fd)
    return tmp, True


@mcp.tool()
async def word_insert_image(path: str = "", data_base64: str = "", url: str = "", position: str = "end",
                            paragraph: int = 0, anchor: str = "", expect: str = "", width_cm: float | None = None,
                            alt_text: str = "", caption: str = "", replace: int = 0, document: str = "") -> dict:
    """Put a picture in the document: a file `path` (png, jpg, gif, bmp, svg, emf - e.g. a screenshot or a
    draw.io export), `data_base64` (raw or data: URI) or an https `url`.
    Placed as its own centered paragraph at position "end" | "start" | "after" / "before" (paragraph N or
    `anchor`) | "section_end" (end of the section under heading `anchor`) | "cursor".
    replace=n swaps embedded image n (from word_context) in place, sized to fit the old picture's box
    (an updated screenshot keeps its size) unless width_cm is given.
    Without width_cm the picture keeps its natural size, fitted to the text width.
    caption adds a numbered "Figure N: caption" line under it."""
    def fn(app):
        d = get_doc(app, document)
        src, temp = _image_file(path, data_base64, url)
        try:
            with undo_step(app, "insert image"):
                box = None
                if replace:
                    meta = get_context(d).images
                    if not 1 <= replace <= len(meta):
                        raise WordError(f"image {replace} does not exist (document has {len(meta)})")
                    k = meta[replace - 1]["inline_index"]
                    if not k:
                        raise WordError(f"image {replace} is a floating shape; only inline images can be replaced")
                    old = d.InlineShapes(k)
                    rng, box = old.Range, (old.Width, old.Height)
                    alt = alt_text or old.AlternativeText
                    old.Delete()
                else:
                    alt = alt_text
                    pos = new_paragraph_at(d, app, position, paragraph, anchor, expect)
                    rng = d.Range(pos, pos)
                    rng.ParagraphFormat.Alignment = ALIGN["center"]
                shp = d.InlineShapes.AddPicture(src, False, True, rng)
                max_w = usable_width_cm(d, shp.Range.Start) * CM_TO_PT
                if width_cm:
                    target = width_cm * CM_TO_PT
                elif box and shp.Width > 0 and shp.Height > 0:
                    # fit inside the old picture's box: same size for an updated screenshot, no blow-up
                    # when the new picture has another shape
                    target = shp.Width * min(box[0] / shp.Width, box[1] / shp.Height)
                else:
                    target = min(shp.Width, max_w)
                if shp.Width > 0 and abs(shp.Width - target) > 0.5:
                    ratio = shp.Height / shp.Width
                    shp.Width = target
                    shp.Height = target * ratio
                if alt:
                    shp.AlternativeText = alt
                if caption:
                    shp.Range.InsertCaption("Figure", ": " + caption, "", 1)  # wdCaptionPositionBelow
                start = shp.Range.Start
        finally:
            if temp:
                try:
                    os.remove(src)
                except OSError:
                    pass
        return {"paragraph": para_index(d, start), "width_cm": round(shp.Width / CM_TO_PT, 2),
                "height_cm": round(shp.Height / CM_TO_PT, 2), "replaced": replace or None,
                "caption": bool(caption)}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_insert_toc(position: str = "start", paragraph: int = 0, anchor: str = "", levels: int = 3,
                          document: str = "") -> dict:
    """Insert a table of contents built from the Heading styles (levels 1..levels), e.g. after the title:
    position="after", anchor="<title text>". Refresh it later with word_update_fields."""
    def fn(app):
        d = get_doc(app, document)
        with undo_step(app, "table of contents", freeze=True):
            pos = new_paragraph_at(d, app, position, paragraph, anchor)
            toc = d.TablesOfContents.Add(d.Range(pos, pos), True, 1, max(1, min(levels, 9)))
            toc.Update()
        return {"tables_of_contents": d.TablesOfContents.Count, "paragraph": para_index(d, pos)}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_update_fields(document: str = "") -> dict:
    """Refresh every field: tables of contents, figure/caption numbers, cross-references, page numbers."""
    def fn(app):
        d = get_doc(app, document)
        with undo_step(app, "update fields", freeze=True):
            d.Fields.Update()
            for i in range(1, d.TablesOfContents.Count + 1):
                d.TablesOfContents(i).Update()
        return {"fields": d.Fields.Count, "tables_of_contents": d.TablesOfContents.Count}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_autosave(on: bool | None = None, document: str = "") -> dict:
    """Read (on=null) or set AutoSave for a SharePoint / OneDrive document. With AutoSave off, edits stay in
    Word until the user saves (Ctrl+S). Local files have no AutoSave (null)."""
    def fn(app):
        d = get_doc(app, document)
        if on is not None:
            try:
                d.AutoSaveOn = on
            except pywintypes.com_error as e:
                raise WordError("AutoSave cannot be changed for this document (local file or policy): "
                                + com_message(e)) from None
        return {"document": d.Name, "autosave": autosave_state(d), "saved": bool(d.Saved)}
    return await worker.run(fn)


@mcp.tool()
async def word_batch(calls: list[dict], stop_on_error: bool = True) -> dict:
    """Run several word_* tools in ONE request (saves a model round trip per call):
    calls = [{"tool": "word_insert_markdown", "args": {...}}, {"tool": "word_format", "args": {...}}, ...].
    Runs in order; image results are summarised as text."""
    results = []
    for i, c in enumerate(calls):
        name = c.get("tool", "")
        fn = globals().get(name)
        if not name.startswith("word_") or name == "word_batch" or not callable(fn):
            results.append({"tool": name, "ok": False, "error": "unknown tool"})
            if stop_on_error:
                break
            continue
        try:
            r = await fn(**(c.get("args") or {}))
            if isinstance(r, list):
                r = [x if isinstance(x, (str, dict)) else "<image>" for x in r]
            results.append({"tool": name, "ok": True, "result": r})
        except Exception as e:  # noqa: BLE001 - reported per call
            results.append({"tool": name, "ok": False, "error": str(e)})
            if stop_on_error:
                break
    return {"ran": len(results), "of": len(calls), "results": results}


@mcp.tool()
async def word_track_changes(action: str = "status", document: str = "") -> dict:
    """Track changes. action: "on" | "off" | "status" | "list" | "accept_all" | "reject_all"."""
    def fn(app):
        d = get_doc(app, document)
        if action in ("on", "off"):
            d.TrackRevisions = action == "on"
        elif action == "accept_all":
            d.Revisions.AcceptAll()
        elif action == "reject_all":
            d.Revisions.RejectAll()
        elif action not in ("status", "list"):
            raise WordError("action must be on, off, status, list, accept_all or reject_all")
        out = {"track_changes": bool(d.TrackRevisions), "revisions": d.Revisions.Count}
        if action == "list":
            kinds = {1: "insert", 2: "delete", 3: "property", 4: "paragraph_number", 5: "display_field",
                     6: "reconcile", 7: "conflict", 8: "style", 9: "replace", 10: "paragraph_property",
                     11: "table_property", 12: "section_property", 13: "style_definition", 14: "moved_from",
                     15: "moved_to", 16: "cell_insertion", 17: "cell_deletion", 18: "cell_merge"}
            items = []
            for i in range(1, min(d.Revisions.Count, COMMENT_LIMIT) + 1):
                r = d.Revisions(i)
                items.append({"n": i, "type": kinds.get(r.Type, r.Type), "author": r.Author,
                              "text": clean(r.Range.Text), "paragraph": para_index(d, r.Range.Start)})
            out["list"] = items
        return out
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_header_footer(header: str | None = None, footer: str | None = None, page_numbers: bool = False,
                             document: str = "") -> dict:
    """Set the primary header and/or footer text of every section; page_numbers adds a centered page number
    to the footer. Leave header/footer out (null) to only read them."""
    def fn(app):
        d = get_doc(app, document)
        with undo_step(app, "header/footer"):
            for i in range(1, d.Sections.Count + 1):
                s = d.Sections(i)
                if header is not None:
                    s.Headers(1).Range.Text = header
                if footer is not None:
                    s.Footers(1).Range.Text = footer
                if page_numbers:
                    s.Footers(1).PageNumbers.Add(1, True)  # wdAlignPageNumberCenter
        s1 = d.Sections(1)
        return {"sections": d.Sections.Count, "header": clean(s1.Headers(1).Range.Text).strip(),
                "footer": clean(s1.Footers(1).Range.Text).strip()}
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_page_setup(orientation: str = "", margins_cm: float | None = None, paper: str = "",
                          document: str = "") -> dict:
    """Page layout for the whole document: orientation (portrait/landscape), margins_cm (all four),
    paper (a4/letter/legal). No arguments returns the current setup."""
    def fn(app):
        d = get_doc(app, document)
        ps = d.PageSetup
        with undo_step(app, "page setup"):
            if orientation:
                ps.Orientation = 1 if orientation.lower() == "landscape" else 0
            if margins_cm is not None:
                # Application.CentimetersToPoints fails through late-bound COM; same math here.
                pts = margins_cm * CM_TO_PT
                ps.TopMargin = ps.BottomMargin = ps.LeftMargin = ps.RightMargin = pts
            if paper:
                sizes = {"a4": 7, "letter": 2, "legal": 4}
                if paper.lower() not in sizes:
                    raise WordError("paper must be a4, letter or legal")
                ps.PaperSize = sizes[paper.lower()]
        return {"orientation": "landscape" if ps.Orientation == 1 else "portrait",
                "page_width_cm": round(ps.PageWidth / CM_TO_PT, 2),
                "page_height_cm": round(ps.PageHeight / CM_TO_PT, 2),
                "margins_cm": {k: round(getattr(ps, k + "Margin") / CM_TO_PT, 2)
                               for k in ("Top", "Bottom", "Left", "Right")},
                "pages": d.ComputeStatistics(2)}  # wdStatisticPages
    return await worker.run(fn, edit=True)


@mcp.tool()
async def word_select(paragraph: int | None = None, find: str = "", occurrence: int = 1, start: int | None = None,
                      end: int | None = None, document: str = "") -> dict:
    """Move Word's cursor/selection to a paragraph, the Nth match of `find`, or start..end, and scroll to it."""
    def fn(app):
        d = get_doc(app, document)
        a, b = targets(d, app, paragraph, None, find, occurrence if find else 0, False, start, end, False)[0]
        d.Activate()
        app.Selection.SetRange(a, b)
        app.ActiveWindow.ScrollIntoView(app.Selection.Range, True)
        return {"start": a, "end": b, "paragraph": para_index(d, a)}
    return await worker.run(fn)


@mcp.tool()
async def word_undo(times: int = 1, redo: bool = False, document: str = "") -> dict:
    """Undo (or redo) the last N edits. Each MCP edit tool is one step."""
    def fn(app):
        d = get_doc(app, document)
        ok = d.Redo(times) if redo else d.Undo(times)
        return {"done": bool(ok), "paragraph_count": d.Paragraphs.Count}
    return await worker.run(fn, edit=True)


_SEG = re.compile(r"\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?:\((.*)\))?\s*$", re.S)


def _split_path(path: str) -> list[str]:
    parts, depth, cur = [], 0, ""
    for ch in path:
        if ch == "." and depth == 0:
            parts.append(cur)
            cur = ""
            continue
        depth += ch == "("
        depth -= ch == ")"
        cur += ch
    parts.append(cur)
    return [p for p in parts if p.strip()]


def _walk(obj, seg: str):
    m = _SEG.match(seg)
    if not m:
        raise WordError(f"bad path segment {seg!r}")
    name, args = m.group(1), m.group(2)
    val = getattr(obj, name)
    if args is not None:
        a = json.loads("[" + args + "]") if args.strip() else []
        val = val(*a)
    return val


@mcp.tool()
async def word_com(path: str, action: str = "get", value: Any = None, args: list | None = None) -> Any:
    """Escape hatch to the full Word object model (see word_api_search). `path` starts at the Application,
    e.g. "ActiveDocument.Paragraphs(3).Range.Font.Size", "ActiveDocument.Sections(1).PageSetup.Orientation",
    "Selection". Arguments in parentheses are JSON values.
    action: "get" reads the property, "set" assigns `value`, "call" invokes the last segment as a method with `args`."""
    def fn(app):
        parts = _split_path(path)
        if not parts:
            raise WordError("empty path")
        obj = app
        for seg in parts[:-1]:
            obj = _walk(obj, seg)
        last = parts[-1]
        if action == "get":
            return py_value(_walk(obj, last))
        if action == "set":
            m = _SEG.match(last)
            if not m or m.group(2) is not None:
                raise WordError("set needs a plain property name as the last segment")
            with undo_step(app, "set " + m.group(1)):
                setattr(obj, m.group(1), value)
            return {"set": path, "value": py_value(getattr(obj, m.group(1)))}
        if action == "call":
            m = _SEG.match(last)
            if not m:
                raise WordError(f"bad method segment {last!r}")
            with undo_step(app, m.group(1)):
                return py_value(getattr(obj, m.group(1))(*(args or [])))
        raise WordError("action must be get, set or call")
    return await worker.run(fn, edit=action != "get")


_api_cache: list | None = None


def build_api_json() -> None:
    """First use: decode this PC's Word type library (MSWORD.OLB) into word-api.json."""
    import subprocess
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                        os.path.join(HERE, "dump_typelib.ps1"), "-OutJson", API_JSON],
                       capture_output=True, text=True, timeout=300)
    if r.returncode != 0 or not os.path.exists(API_JSON):
        raise WordError("could not build word-api.json: " + (r.stderr or r.stdout).strip()[:500])


@mcp.tool()
async def word_api_search(query: str, max_results: int = 40) -> dict:
    """Search the decoded Word object model (MSWORD.OLB): object names, members with parameter lists,
    and enum constants. Query "Range" lists Range's members; "Range.Find" or "InsertBreak" finds members;
    "WdBreakType" lists an enum's values."""
    global _api_cache
    if _api_cache is None:
        if not os.path.exists(API_JSON):
            await asyncio.to_thread(build_api_json)
        with open(API_JSON, encoding="utf-8") as fh:
            _api_cache = json.load(fh)
    q = query.strip()
    obj_q, _, mem_q = q.partition(".") if "." in q else ("", "", q)
    hidden = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke"}

    def sig(m):
        if m["kind"] == "const":
            return f"{m['name']} = {m['value']}"
        ps = ", ".join(("[" + p["name"] + "]" if p["optional"] else p["name"]) + ": " + p["type"] for p in m["params"])
        return f"{m['kind']} {m['name']}({ps}) -> {m['returns']}"

    out = []
    for t in _api_cache:
        tname = t["name"].lstrip("_")
        exact_type = q.lower() in (t["name"].lower(), tname.lower())
        if exact_type:
            members = [m for m in t["members"] if not m.get("hidden") and m["name"] not in hidden]
            return {"type": t["name"], "kind": t["kind"], "members": [sig(m) for m in members]}
        if obj_q and obj_q.lower() not in (t["name"].lower(), tname.lower()):
            continue
        for m in t["members"]:
            if m.get("hidden") or m["name"] in hidden:
                continue
            if mem_q.lower() in m["name"].lower():
                out.append(f"{t['name']}: {sig(m)}")
                if len(out) >= max_results:
                    return {"matches": out, "truncated": True}
    return {"matches": out, "truncated": False}


if __name__ == "__main__":
    mcp.run()
