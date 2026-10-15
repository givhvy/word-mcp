"""End-to-end test: launches word_mcp_server.py over stdio, runs every tool against a
scratch document in the running Word, checks results and prints per-call latency."""

import asyncio
import json
import os
import shutil
import sys
import time

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # python -I leaves this folder off sys.path
import tests_common  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "test-output")
timings: list[tuple[str, float]] = []
failures: list[str] = []


async def call(s: ClientSession, tool: str, label: str = "", **args):
    t = time.perf_counter()
    r = await s.call_tool(tool, args)
    ms = (time.perf_counter() - t) * 1000
    timings.append((label or tool, ms))
    text = "".join(c.text for c in r.content if hasattr(c, "text"))
    if r.isError:
        raise AssertionError(f"{tool} failed: {text}")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


async def main():
    tests_common.setup()
    os.makedirs(OUT, exist_ok=True)
    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    t0 = time.perf_counter()
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        tools = (await s.list_tools()).tools
        print(f"server up in {(time.perf_counter() - t0) * 1000:.0f} ms, {len(tools)} tools")

        st = await call(s, "word_status", "status (first call, attaches to Word)")
        print(f"Word {st['build']}, {len(st['documents'])} open document(s)")
        user_docs = {d["name"] for d in st["documents"]}

        d = await call(s, "word_new_document", text="Intro paragraph about the quarterly plan.\nSecond line with a typo: teh budget.")
        doc = d["name"]
        print("scratch document:", doc)
        check(d["paragraphs"] == 2, "new document has 2 paragraphs")
        try:
            await scenario(s, doc, user_docs)
        finally:
            now = await s.call_tool("word_status", {})
            for x in json.loads(now.content[0].text)["documents"]:
                if x["name"] not in user_docs:  # only ever close what the test created
                    await s.call_tool("word_close", {"document": x["name"]})
        end = await call(s, "word_status", "status (final)")
        now_docs = {x["name"] for x in end["documents"]}
        check(now_docs == user_docs, "user's documents untouched, scratch docs closed - "
              + tests_common.doc_diff(user_docs, now_docs))
    report()


async def scenario(s: ClientSession, doc: str, user_docs: set):
    if True:

        await call(s, "word_insert", "insert title at start", text="Quarterly Plan", position="start", style="Title", document=doc)
        await call(s, "word_insert", "insert heading at end", text="Goals", style="Heading 1", document=doc)
        ins = await call(s, "word_insert", "insert 3 bullets", text="Ship v2\nHire two engineers\nCut costs 10%",
                         style="List Bullet", document=doc)
        check(ins["last_paragraph"] - ins["first_paragraph"] == 2, "3 bullet paragraphs inserted")
        await call(s, "word_insert", "insert after paragraph 1", text="Prepared by the team.", position="after", paragraph=1, document=doc)

        rd = await call(s, "word_read", "read with styles", with_styles=True, document=doc)
        texts = [p["text"] for p in rd["paragraphs"]]
        styles = [p.get("style") for p in rd["paragraphs"]]
        print("  paragraphs:", json.dumps([(p["n"], p["text"][:30], p.get("style")) for p in rd["paragraphs"]]))
        check(texts[0] == "Quarterly Plan" and styles[0] == "Title", "title first with Title style")
        check(texts[1] == "Prepared by the team.", "insert after paragraph 1 landed at 2")
        check(texts[4] == "Goals" and styles[4] == "Heading 1", "heading 1 applied")
        check(texts[5:8] == ["Ship v2", "Hire two engineers", "Cut costs 10%"] and styles[5] == "List Bullet", "bullets with List Bullet")
        check(rd["paragraph_count"] == len(rd["paragraphs"]), "read count matches Word paragraph count")

        f = await call(s, "word_find", text="teh", document=doc)
        check(f["count"] == 1 and f["matches"][0]["paragraph"] == 4, "find 'teh' in paragraph 4")
        rp = await call(s, "word_replace", find="teh", replace="the", document=doc)
        check(rp["count"] == 1, "replace fixed 1 typo")

        await call(s, "word_format", "format find: bold+highlight", find="budget", bold=True, highlight="yellow", document=doc)
        await call(s, "word_format", "format paragraph: center+color", paragraph=1, align="center", color="#1F4E79", document=doc)
        bold = await call(s, "word_com", "com get Bold", path=f'Documents("{doc}").Range({f["matches"][0]["start"] + 4}, {f["matches"][0]["start"] + 10}).Font.Bold')
        check(bold == -1, "'budget' is bold (Word true = -1)")

        tb = await call(s, "word_insert_table", rows=[["Item", "Owner", "Due"], ["Ship v2", "Ana", "Q3"],
                                                      ["Hire", "Ben", "Q4"]], after_paragraph=4, document=doc)
        check(tb["table"] == 1 and tb["rows"] == 3 and tb["columns"] == 3, "3x3 table inserted")
        await call(s, "word_set_cell", table=1, row=3, column=3, text="Q1 next year", document=doc)
        rt = await call(s, "word_read_tables", document=doc)
        cells = rt["tables"][0]["cells"]
        check(cells[0] == ["Item", "Owner", "Due"] and cells[2][2] == "Q1 next year", "table read back + cell edit")
        rd2 = await call(s, "word_read", "read after table", document=doc)
        check(rd2["paragraph_count"] == len(rd2["paragraphs"]), "read count still matches with a table")

        await call(s, "word_add_comment", comment="Is Ben free?", find="Ben", document=doc)
        lc = await call(s, "word_list_comments", document=doc)
        check(lc["comment_count"] == 1 and lc["comments"][0]["on"] == "Ben", "comment on 'Ben'")

        await call(s, "word_track_changes", "track changes on", action="on", document=doc)
        gp = next(p["n"] for p in rd2["paragraphs"] if p["text"] == "Goals")
        await call(s, "word_set_paragraph", paragraph=gp, text="Goals for Q3", document=doc)
        tc = await call(s, "word_track_changes", "list revisions", action="list", document=doc)
        check(tc["revisions"] >= 1, f"tracked revision recorded ({tc['revisions']})")
        await call(s, "word_track_changes", "accept all", action="accept_all", document=doc)
        tc2 = await call(s, "word_track_changes", "track changes off", action="off", document=doc)
        check(tc2["revisions"] == 0 and not tc2["track_changes"], "revisions accepted, tracking off")

        hf = await call(s, "word_header_footer", header="Quarterly Plan - Draft", footer="Confidential", page_numbers=True, document=doc)
        check(hf["header"] == "Quarterly Plan - Draft", "header set")
        ps = await call(s, "word_page_setup", orientation="portrait", margins_cm=2.0, paper="a4", document=doc)
        check(abs(ps["margins_cm"]["Top"] - 2.0) < 0.05 and abs(ps["page_width_cm"] - 21.0) < 0.1, "A4, 2 cm margins")

        sel = await call(s, "word_select", find="Hire two engineers", document=doc)
        stt = await call(s, "word_status", "status (selection)")
        check(stt["active"]["selection"]["text"] == "Hire two engineers", "selection moved to the found text")

        api = await call(s, "word_api_search", query="WdBreakType")
        check(any("wdPageBreak" in m for m in api["members"]), "API search lists WdBreakType.wdPageBreak")
        api2 = await call(s, "word_api_search", "api search member", query="InsertBreak")
        check(len(api2["matches"]) > 0, "API search finds InsertBreak")
        await call(s, "word_com", "com call InsertBreak", path=f'Documents("{doc}").Paragraphs(1).Range.InsertParagraphAfter', action="call")

        n0 = (await call(s, "word_status", "status"))["active"]["paragraphs"]
        await call(s, "word_undo", document=doc)
        n1 = (await call(s, "word_status", "status"))["active"]["paragraphs"]
        check(n1 == n0 - 1, "undo removed the paragraph added by word_com")

        # Bulk / speed
        big = "\n".join(f"Line {i}: the quick brown fox jumps over the lazy dog." for i in range(1, 501))
        await call(s, "word_insert", "BULK insert 500 paragraphs", text=big, document=doc)
        txt = await call(s, "word_get_text", "BULK get_text whole doc", document=doc)
        check("Line 500:" in txt, "500 lines present")
        rb = await call(s, "word_read", "BULK read all paragraphs", document=doc)
        check(rb["paragraph_count"] == len(rb["paragraphs"]), f"bulk read {len(rb['paragraphs'])} paragraphs, count matches")
        rr = await call(s, "word_replace", "BULK replace 500 occurrences", find="lazy dog", replace="sleepy cat", document=doc)
        check(rr["count"] == 500, f"replaced {rr['count']} of 500")
        fm = await call(s, "word_format", "BULK format 500 matches", find="sleepy cat", italic=True, highlight="green", document=doc)
        check(fm["formatted_matches"] == 500, "formatted 500 matches")
        it = await call(s, "word_com", "com get Italic", path=f'Documents("{doc}").Paragraphs(30).Range.Font.Italic')
        check(it == 9999999, "paragraph 30 is partly italic (mixed = 9999999)")
        bigt = [["#", "Name", "Value", "Note"]] + [[str(i), f"row {i}", str(i * 3), "ok"] for i in range(1, 101)]
        await call(s, "word_insert_table", "BULK table 101x4", rows=bigt, document=doc)
        await call(s, "word_delete_paragraphs", "delete paragraphs", start_paragraph=3, document=doc)

        sv = await call(s, "word_save", "save .docx", path=os.path.join(OUT, "mcp-test.docx"), document="")
        check(sv["size_bytes"] > 0, "saved docx")
        pdf = await call(s, "word_save", "export .pdf", path=os.path.join(OUT, "mcp-test.pdf"), document=sv["saved"].split(os.sep)[-1])
        check(pdf["size_bytes"] > 0, f"exported pdf ({pdf['size_bytes']} bytes)")
        shutil.copyfile(os.path.join(OUT, "mcp-test.docx"), os.path.join(OUT, "mcp-test-copy.docx"))
        op = await call(s, "word_open", path=os.path.join(OUT, "mcp-test-copy.docx"), read_only=True)
        check(op["paragraphs"] > 500 and "replaced_blank_document" not in op, "opened saved copy, nothing replaced")
        await call(s, "word_close", "close copy", document="mcp-test-copy.docx")
        await call(s, "word_close", document="mcp-test.docx")

        try:
            await call(s, "word_read", "error path", document="does-not-exist.docx")
            check(False, "error for unknown document")
        except AssertionError as e:
            msg = str(e)
            check("No open document named" in msg or "No document is open" in msg, "clean error for unknown document")


def report():
    print("\nlatency per call (ms):")
    for label, ms in timings:
        print(f"  {ms:8.1f}  {label}")
    small = [ms for label, ms in timings if not label.startswith(("BULK", "status (first", "save", "export", "word_open", "word_new"))]
    small.sort()
    print(f"\n{len(timings)} calls; median {small[len(small) // 2]:.1f} ms, p90 {small[int(len(small) * .9)]:.1f} ms for regular tools")
    print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
    sys.exit(1 if failures else 0)


asyncio.run(main())
