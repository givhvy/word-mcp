"""Regression tests for the anchor / numbering bugs reported against d35b0d4: in a document with
section breaks, layout and nested tables, an overview and a references table repeating provider names,
a TOC, images with captions, comments, hidden text and a content control, every insert must land
exactly where asked (or be refused), deletes must remove only what they say, and word_read /
expect / word_changes must use Word's own paragraph numbers."""

import asyncio
import json
import os
import re
import sys

import pythoncom
import win32com.client
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # run with python -I, which leaves the script folder off sys.path
import tests_common  # noqa: E402
import tests_fixtures  # noqa: E402

OUT = os.path.join(HERE, "test-output")
failures: list[str] = []
WD_IN_TABLE, WD_SECTION = 12, 2  # Range.Information(wdWithInTable / wdActiveEndSectionNumber)


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


async def call(s, tool, **args):
    r = await s.call_tool(tool, args)
    text = "".join(c.text for c in r.content if c.type == "text")
    if r.isError:
        raise AssertionError(text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


async def refused(s, tool, needle, **args):
    try:
        await call(s, tool, **args)
        return False, "not refused"
    except AssertionError as e:
        return needle.lower() in str(e).lower(), str(e)[:220]


def alnum(t):
    return re.sub(r"[\W_]+", "", t.lower())


class Doc:
    """Ground truth straight from Word (never through the MCP)."""

    def __init__(self, d):
        self.d = d

    def index_of(self, text):
        hits = [i for i in range(1, self.d.Paragraphs.Count + 1)
                if alnum(self.d.Paragraphs(i).Range.Text).startswith(alnum(text))]
        return hits

    def text(self, i):
        return self.d.Paragraphs(i).Range.Text.strip("\r\x07\x0c ")

    def info(self, i, what):
        return self.d.Paragraphs(i).Range.Information(what)


async def main():
    tests_common.setup()
    pythoncom.CoInitialize()
    app = win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject("Word.Application")._oleobj_)
    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        before = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        d = tests_fixtures.build_integration_doc(app, OUT)
        doc, W = d.Name, Doc(d)
        D = {"document": doc}
        try:
            sections0, images0, comments0 = d.Sections.Count, d.InlineShapes.Count, d.Comments.Count
            # --- numbering: word_read == Word, paragraph by paragraph -----------------------------
            rd = await call(s, "word_read", **D)
            bad = [p["n"] for p in rd["paragraphs"]
                   if alnum(p["text"]) != alnum(W.text(p["n"])) and alnum(W.text(p["n"])) not in ("",)]
            check(rd["paragraph_count"] == d.Paragraphs.Count and not bad,
                  f"word_read numbering == Word's for all {d.Paragraphs.Count} paragraphs (mismatches: {bad[:5]})")
            ctx = (await s.call_tool("word_context", D)).content[0].text
            check("paragraph numbers may drift" not in ctx, "word_context numbering matches Word")

            # --- 1. section_end on a heading whose name also sits in two tables -------------------
            res = await call(s, "word_insert_markdown", markdown="MARK-SECTION-END: VeloViewer details.",
                             position="section_end", anchor="VeloViewer", **D)
            [i] = W.index_of("MARK-SECTION-END")
            vv = W.index_of("VeloViewer provides")[0]
            gsv = [k for k in W.index_of("Google Street View") if not W.info(k, WD_IN_TABLE)
                   and W.d.Paragraphs(k).Style.NameLocal.lower().startswith("heading")][0]
            check(vv < i < gsv, f"section_end lands inside the VeloViewer section (p{vv} < p{i} < p{gsv})")
            check(W.info(i, WD_SECTION) == W.info(vv, WD_SECTION),
                  "section_end stays before the section break (same page section as VeloViewer)")
            check(not W.info(i, WD_IN_TABLE), "section_end is not in a table")
            print(f"    result placed_after={res.get('placed_after')!r} placed_before={res.get('placed_before')!r}")

            # --- 2. before the Google Street View heading (also in tables and the TOC) -------------
            await call(s, "word_insert_markdown", markdown="MARK-BEFORE-GSV.", position="before",
                       anchor="Google Street View", **D)
            [i] = W.index_of("MARK-BEFORE-GSV")
            check(not W.info(i, WD_IN_TABLE) and W.text(i + 1) == "Google Street View"
                  and W.d.Paragraphs(i + 1).Style.NameLocal.lower().startswith("heading"),
                  "before anchor lands right before the heading, not in the overview table")

            # --- 3. after a caption -------------------------------------------------------------------
            cap = "Figure 1. VeloViewer integration flow."
            await call(s, "word_insert_markdown", markdown="MARK-AFTER-CAPTION.", position="after", anchor=cap, **D)
            [i] = W.index_of("MARK-AFTER-CAPTION")
            check(alnum(W.text(i - 1)) == alnum(cap) and W.text(i + 1) == "How it is called / request body",
                  f"after caption lands right after it (p{i - 1} {W.text(i - 1)[:30]!r} / p{i + 1} {W.text(i + 1)[:30]!r})")

            # --- 4. paragraph + expect, numbers taken from word_read ------------------------------------
            rd = await call(s, "word_read", **D)
            n_cap = next(p["n"] for p in rd["paragraphs"] if alnum(p["text"]) == alnum(cap))
            check(alnum(W.text(n_cap)) == alnum(cap), f"word_read's caption number p{n_cap} is Word's caption")
            await call(s, "word_insert", text="MARK-PARA-EXPECT.", position="after", paragraph=n_cap, expect=cap, **D)
            [i] = W.index_of("MARK-PARA-EXPECT")
            check(i == n_cap + 1, "paragraph+expect insert lands right after the caption")
            ok, msg = await refused(s, "word_insert", "document changed", text="x", position="after",
                                    paragraph=n_cap, expect="Google Street View protocol", **D)
            check(ok, f"wrong expect refused ({msg[:90]})")

            # --- ambiguous / table-only anchors are refused, never guessed ------------------------------
            ok, msg = await refused(s, "word_insert", "ambiguous", text="x", position="after", anchor="integration flow", **D)
            check(ok, f"ambiguous anchor refused with candidates ({msg[:120]})")
            ok, msg = await refused(s, "word_insert", "table", text="x", position="after", anchor="Ride analytics", **D)
            check(ok, f"table-only anchor refused with a hint ({msg[:120]})")
            ok, msg = await refused(s, "word_insert_markdown", "heading", markdown="x", position="section_end",
                                    anchor="Ride segments and power data", **D)
            check(ok, "section_end on a non-heading refused")
            await call(s, "word_insert", text="MARK-IN-TABLE", position="after", anchor="Ride analytics", in_tables=True, **D)
            [i] = W.index_of("MARK-IN-TABLE")
            check(bool(W.info(i, WD_IN_TABLE)), "in_tables=true targets the table cell on purpose")

            check(d.Sections.Count == sections0, f"section breaks preserved ({d.Sections.Count} sections)")
            check(d.InlineShapes.Count == images0 and d.Comments.Count == comments0, "no images or comments lost by inserts")

            # --- word_changes numbers after section breaks ------------------------------------------------
            await call(s, "word_context", **D)
            k = [k for k in W.index_of("Requests are signed")][0]  # section 3, after two section breaks
            r0 = d.Paragraphs(k).Range
            d.Range(r0.Start, r0.Start).InsertBefore("USER-LINE in section three\r")  # the user types
            ch = await call(s, "word_changes", **D)
            [u] = W.index_of("USER-LINE")
            added = [c for c in ch["changes"] if c["type"] == "added"]
            check(added and added[0].get("paragraphs") == str(u),
                  f"word_changes reports Word's number for the user's line (got {added and added[0].get('paragraphs')}, Word p{u})")

            # --- 2. deletes remove exactly what they say ----------------------------------------------------
            res = await call(s, "word_delete_paragraphs", anchor="MARK-AFTER-CAPTION.", expect="MARK-AFTER-CAPTION.", **D)
            check(res["deleted"] == 1 and res["removed"] == ["MARK-AFTER-CAPTION."] and not any(res["removed_content"].values()),
                  f"delete removes only the named paragraph ({res['removed']})")
            ok, msg = await refused(s, "word_delete_paragraphs", "no paragraph matches", anchor="MARK-AFTER-CAPTION.", **D)
            check(ok, "repeat delete reports the anchor is gone")
            check(d.InlineShapes.Count == images0 and d.Comments.Count == comments0, "delete lost no images or comments")
            img_para = next(k for k in range(1, d.Paragraphs.Count + 1) if d.Paragraphs(k).Range.InlineShapes.Count)
            ok, msg = await refused(s, "word_delete_paragraphs", "picture", start_paragraph=img_para,
                                    end_paragraph=img_para + 1, **D)
            check(ok, f"deleting an image paragraph is refused without allow_content_loss ({msg[:80]})")
            brk = next(k for k in range(1, d.Paragraphs.Count) if d.Paragraphs(k).Range.Text.endswith("\x0c"))
            ok, msg = await refused(s, "word_delete_paragraphs", "section break", start_paragraph=brk, end_paragraph=brk, **D)
            check(ok, "deleting a section break is refused without allow_content_loss")
            ok, msg = await refused(s, "word_set_paragraph", "comment", anchor="Google Street View protocol", text="x", **D)
            check(ok, "replacing a commented paragraph's text is refused")
            check(d.Sections.Count == sections0 and d.InlineShapes.Count == images0 and d.Comments.Count == comments0,
                  "after all refusals: nothing lost")

            # --- replace a section between two section breaks -------------------------------------------
            rs = await call(s, "word_replace_section", anchor="VeloViewer", markdown="REWRITTEN VeloViewer section.", **D)
            vv_h = [k for k in W.index_of("VeloViewer") if d.Paragraphs(k).Style.NameLocal.lower().startswith("heading")][0]
            check(W.text(vv_h + 1) == "REWRITTEN VeloViewer section." and d.Sections.Count == sections0,
                  f"replace_section keeps heading and section breaks (removed {rs['removed_paragraphs']}, "
                  f"content {rs['removed_content']}, layout {rs.get('layout_kept')})")
            check(W.index_of("MARK-BEFORE-GSV"), "text placed after the section break (another page section) is kept")
            check(W.info(vv_h + 1, WD_SECTION) == W.info(vv_h, WD_SECTION), "rewritten text stays in its section")
            check(W.index_of("Google Street View protocol"), "next section untouched")

            # --- 4. status while a comment is selected in the comments pane --------------------------------
            try:
                d.ActiveWindow.View.SplitSpecial = 15  # wdPaneComments
                d.Comments(1).Range.Select()
            except pythoncom.com_error as e:
                print("    (could not open the comments pane:", e, ")")
            st = await call(s, "word_status")
            check(isinstance(st, dict) and st.get("documents"), "word_status works while a comment is selected")
            try:
                d.ActiveWindow.View.SplitSpecial = 0
            except pythoncom.com_error:
                pass
        finally:
            for name in [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]:
                if name not in before:
                    app.Documents(name).Close(0)
        after = {x["name"] for x in (await call(s, "word_status"))["documents"]}
        check(after == before, "only the test's documents were closed - " + tests_common.doc_diff(before, after))
    print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
    sys.exit(1 if failures else 0)


asyncio.run(main())
