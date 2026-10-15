"""Tests for word_insert_table placement and every word_table_edit action, checked against Word's real
state over COM, on a disposable fixture document."""

import asyncio
import json
import os
import sys

import pythoncom
import win32com.client
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # python -I leaves this folder off sys.path
import tests_common  # noqa: E402
import tests_fixtures  # noqa: E402

OUT = os.path.join(HERE, "test-output")
failures: list[str] = []


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


async def call(s, tool, **args):
    r = await s.call_tool(tool, args)
    text = "".join(c.text for c in r.content if c.type == "text")
    if r.isError:
        raise AssertionError(text)
    return json.loads(text)


async def refused(s, tool, needle, **args):
    try:
        await call(s, tool, **args)
        return False, "not refused"
    except AssertionError as e:
        return needle.lower() in str(e).lower(), str(e)[:200]


def cell(t, r, c):
    return t.Cell(r, c).Range.Text[:-2].strip()


async def main():
    tests_common.setup()
    pythoncom.CoInitialize()
    app = win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject("Word.Application")._oleobj_)
    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        before = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        d = tests_fixtures.build_integration_doc(app, OUT)
        D = {"document": d.Name}
        try:
            # --- find tables -----------------------------------------------------------------------
            rt = await call(s, "word_read_tables", near="Strava", **D)
            ov = rt["tables"][0]["table"]
            check(rt["tables"][0]["cells"][0] == ["Provider", "Purpose", "Status"], f"near='Strava' finds the overview table (#{ov})")
            ok, msg = await refused(s, "word_read_tables", "matches tables", near="Google Street View", **D)
            check(ok, f"text in two tables is refused as ambiguous ({msg[:90]})")

            # --- insert by anchor / old form ---------------------------------------------------------
            ins = await call(s, "word_insert_table", rows=[["Metric", "Value"], ["p95", "180"], ["p50", "40"],
                                                           ["p99", "900"]], position="after",
                             anchor="Requests are signed with a key.", **D)
            T = d.Tables(ins["table"])
            prev = d.Range(T.Range.Start - 2, T.Range.Start - 1).Paragraphs(1).Range.Text.strip()
            check(prev == "Requests are signed with a key." and ins["placed_after"].startswith("Requests are signed"),
                  f"table placed right after its anchor (after {ins['placed_after']!r})")
            n_tables = d.Tables.Count
            old = await call(s, "word_insert_table", rows=[["a", "b"]], after_paragraph=1, **D)
            check(d.Tables.Count == n_tables + 1 and old["placed_after"].startswith("Integration Overview"),
                  "after_paragraph (older form) still works")
            await call(s, "word_table_edit", action="delete_table", table=old["table"], **D)
            t_id = next(i for i in range(1, d.Tables.Count + 1) if cell(d.Tables(i), 1, 1) == "Metric")
            E = {"table": t_id, **D}
            T = d.Tables(t_id)

            # --- rows / columns / fill ----------------------------------------------------------------
            res = await call(s, "word_table_edit", action="add_rows", row=2, values=[["p75", "120"], ["p90", "150"]], **E)
            check(T.Rows.Count == 6 and cell(T, 3, 1) == "p75" and cell(T, 4, 1) == "p90" and cell(T, 5, 1) == "p50",
                  "add_rows inserts filled rows after row 2 (count follows values)")
            await call(s, "word_table_edit", action="add_rows", count=1, **E)
            check(T.Rows.Count == 7, "add_rows with no row appends at the end")
            await call(s, "word_table_edit", action="delete_rows", row=7, **E)
            check(T.Rows.Count == 6, "delete_rows removes one row")
            await call(s, "word_table_edit", action="add_columns", values=["Unit", "ms", "ms", "ms", "ms", "ms"], **E)
            check(T.Columns.Count == 3 and cell(T, 1, 3) == "Unit" and cell(T, 6, 3) == "ms", "add_columns appends a filled column")
            await call(s, "word_table_edit", action="fill", row=2, column=3, values=[["s"], ["s"]], **E)
            check(cell(T, 2, 3) == "s" and cell(T, 3, 3) == "s" and cell(T, 4, 3) == "ms", "fill writes a block of cells")
            ok, msg = await refused(s, "word_table_edit", "do not fit", action="fill", row=6, values=[["x"], ["y"]], **E)
            check(ok, "fill outside the table is refused")
            await call(s, "word_table_edit", action="delete_columns", column=3, **E)
            check(T.Columns.Count == 2, "delete_columns removes the column")
            ok, msg = await refused(s, "word_table_edit", "delete_table", action="delete_rows", **E)
            check(ok, "deleting every row points to delete_table")

            # --- formatting ---------------------------------------------------------------------------
            await call(s, "word_table_edit", action="shade", row=1, color="#1F4E79", **E)
            await call(s, "word_table_edit", action="bold", row=1, on=True, **E)
            await call(s, "word_table_edit", action="align", column=2, align="right", **E)
            check(T.Cell(1, 2).Shading.BackgroundPatternColor == 0x794E1F and T.Cell(1, 1).Range.Font.Bold == -1
                  and T.Cell(4, 2).Range.ParagraphFormat.Alignment == 2, "shade / bold header row, right-align a column")
            await call(s, "word_table_edit", action="width", column=1, width_cm=4, **E)
            check(abs(T.Cell(2, 1).Width - 4 * 72 / 2.54) < 2, f"column width 4 cm ({T.Cell(2, 1).Width:.0f} pt)")
            await call(s, "word_table_edit", action="sort", column=2, numeric=True, descending=True, **E)
            vals = [cell(T, i, 2) for i in range(2, T.Rows.Count + 1)]
            check(vals == ["900", "180", "150", "120", "40"] and cell(T, 1, 1) == "Metric",
                  f"numeric descending sort keeps the header on top ({vals})")
            await call(s, "word_table_edit", action="header_rows", count=1, **E)
            check(T.Rows(1).HeadingFormat == -1 and T.Rows(2).HeadingFormat == 0, "header row repeats on each page")
            await call(s, "word_table_edit", action="style", style="Table Grid", **E)
            check(T.Style.NameLocal == "Table Grid", "table style applied")
            await call(s, "word_table_edit", action="borders", on=False, **E)
            check(T.Borders.Enable == 0, "borders off")
            await call(s, "word_table_edit", action="autofit", fit="window", **E)
            check(True, "autofit to window accepted")

            # --- one Ctrl+Z per edit -------------------------------------------------------------------
            rows_now = T.Rows.Count
            await call(s, "word_table_edit", action="add_rows", values=[["undo-me", "1"]], **E)
            await call(s, "word_undo", **D)
            check(d.Tables(t_id).Rows.Count == rows_now, "one undo reverts one table edit")

            # --- merge / split / merged-cell guidance ---------------------------------------------------
            await call(s, "word_table_edit", action="merge", row=2, column=1, end_row=3, end_column=1, **E)
            ok, msg = await refused(s, "word_table_edit", "merged", action="delete_rows", row=4, **E)
            check(ok, f"row operation on a vertically merged table explains the merged cells ({msg[:110]})")
            await call(s, "word_table_edit", action="split", row=2, column=1, split_rows=2, split_columns=1, **E)
            await call(s, "word_table_edit", action="delete_rows", row=4, **E)
            check(True, "split restores row operations")

            # --- protection ------------------------------------------------------------------------------
            T = d.Tables(t_id)
            d.Comments.Add(T.Cell(2, 2).Range, "Is p95 right?")
            ok, msg = await refused(s, "word_table_edit", "comment", action="delete_rows", row=2, **E)
            check(ok, "deleting a commented row is refused")
            res = await call(s, "word_table_edit", action="delete_rows", row=2, allow_content_loss=True, **E)
            check(res["rows"] == T.Rows.Count, "allow_content_loss deletes it")
            pic = os.path.join(OUT, "fx-veloviewer.png")
            d.InlineShapes.AddPicture(pic, False, True, T.Cell(2, 1).Range)
            ok, msg = await refused(s, "word_table_edit", "picture", action="delete_table", **E)
            check(ok, "deleting a table holding a picture is refused")
            n = d.Tables.Count
            await call(s, "word_table_edit", action="delete_table", allow_content_loss=True, **E)
            check(d.Tables.Count == n - 1, "delete_table with allow_content_loss removes it")
            ok, msg = await refused(s, "word_table_edit", "action must be", action="explode", table=1, **D)
            check(ok, "unknown action refused with the list of actions")
        finally:
            for name in [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]:
                if name not in before:
                    app.Documents(name).Close(0)
        after = {x["name"] for x in (await call(s, "word_status"))["documents"]}
        check(after == before, "only the test's documents were closed - " + tests_common.doc_diff(before, after))
    print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
    sys.exit(1 if failures else 0)


asyncio.run(main())
