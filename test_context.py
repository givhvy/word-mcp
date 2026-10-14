"""Tests for comment threads, whole-document context and images (word_mcp_server.py over stdio)."""

import asyncio
import base64
import json
import os
import sys
import time

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "test-output")
failures: list[str] = []
timings: list[tuple[str, float]] = []


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


async def raw(s, tool, label="", **args):
    t = time.perf_counter()
    r = await s.call_tool(tool, args)
    timings.append((label or tool, (time.perf_counter() - t) * 1000))
    if r.isError:
        raise AssertionError(f"{tool}: " + "".join(getattr(c, "text", "") for c in r.content))
    return r


async def call(s, tool, label="", **args):
    r = await raw(s, tool, label, **args)
    text = "".join(c.text for c in r.content if c.type == "text")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


async def main():
    os.makedirs(OUT, exist_ok=True)
    pic = os.path.join(OUT, "test-chart.png")
    im = Image.new("RGB", (480, 240), "white")
    dr = ImageDraw.Draw(im)
    for i, h in enumerate((60, 140, 100, 200)):
        dr.rectangle((40 + i * 100, 220 - h, 110 + i * 100, 220), fill=(40, 90, 170))
    dr.text((40, 10), "Q1-Q4 SALES", fill="black")
    im.save(pic)

    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        names = {t.name for t in (await s.list_tools()).tools}
        print(len(names), "tools")
        before = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        doc = (await call(s, "word_new_document", text="Sales Review"))["name"]
        try:
            await scenario(s, doc, pic)
        finally:
            now = await call(s, "word_status")
            for d in now["documents"]:
                if d["name"] not in before:
                    await s.call_tool("word_close", {"document": d["name"]})
        after = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        check(after == before, "only the test's documents were closed")
    print("\nlatency (ms):")
    for label, ms in timings:
        print(f"  {ms:8.1f}  {label}")
    print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
    sys.exit(1 if failures else 0)


async def scenario(s, doc, pic):
    D = {"document": doc}
    await call(s, "word_format", paragraph=1, style="Title", **D)
    await call(s, "word_insert", text="Summary", style="Heading 1", **D)
    await call(s, "word_insert", text="Revenue grew in every quarter. The fourth quarter was the strongest.", **D)
    await call(s, "word_insert", text="Wins", style="Heading 2", **D)
    await call(s, "word_insert", text="New distributor in Spain\nTwo enterprise renewals", style="List Bullet", **D)
    await call(s, "word_insert", text="Next steps", style="Heading 2", **D)
    await call(s, "word_insert", text="Hire a second sales rep\nLaunch the partner program", style="List Number", **D)
    await call(s, "word_format", find="every quarter", bold=True, **D)
    await call(s, "word_format", find="strongest", italic=True, **D)
    await call(s, "word_insert_table", rows=[["Quarter", "Revenue"], ["Q1", "1.2M"], ["Q2", "1.9M"],
                                            ["Q3", "1.6M"], ["Q4", "2.4M"]], **D)
    img = await call(s, "word_insert_image", path=pic, position="after", paragraph=3, width_cm=8,
                     alt_text="Bar chart of quarterly sales", **D)
    check(img["paragraph"] == 4 and abs(img["width_cm"] - 8) < 0.1, f"image inserted as paragraph 4, 8 cm wide ({img})")

    # Comments created out of document order, then threaded.
    c_late = await call(s, "word_add_comment", comment="Add the Spanish distributor's name.", find="Spain", **D)
    c_early = await call(s, "word_add_comment", comment="Is this before or after refunds?", find="Revenue grew", **D)
    check(c_early["n"] == 1 and c_late["n"] == 1 or c_early["n"] < c_late["n"] + 1,
          f"comment numbers follow document order (early {c_early['n']})")
    rep1 = await call(s, "word_reply_comment", n=c_early["n"], text="After refunds.", **D)
    rep2 = await call(s, "word_reply_comment", n=rep1["n"], text="Thanks, resolving.", **D)
    check(rep2["thread_replies"] == 2, "reply to a reply stays on the same thread")
    lc = await call(s, "word_list_comments", **D)
    threads = lc["comments"]
    print("  threads:", json.dumps([(t["n"], t["text"], [r["text"] for r in t["replies"]]) for t in threads]))
    check(lc["comment_count"] == 4 and lc["threads"] == 2, "4 comments in 2 threads")
    first = threads[0]
    check(first["on"] == "Revenue grew" and [r["text"] for r in first["replies"]] == ["After refunds.", "Thanks, resolving."],
          "first thread is the earlier anchor with both replies in order")

    # n in the list must be Word's own Comments(n): check every one against COM.
    flat = [(t["n"], t["text"]) for t in threads] + [(r["n"], r["text"]) for t in threads for r in t["replies"]]
    ok = True
    for n, text in flat:
        com = await call(s, "word_com", path=f'Documents("{doc}").Comments({n}).Range.Text')
        ok &= com.strip() == text
    check(ok, "every listed n matches Word's Comments(n)")

    await call(s, "word_resolve_comment", n=rep2["n"], **D)
    lc2 = await call(s, "word_list_comments", include_resolved=False, **D)
    check(lc2["open_threads"] == 1 and len(lc2["comments"]) == 1 and lc2["comments"][0]["on"] == "Spain",
          "resolving via a reply resolves its thread; filter hides it")
    late_n = lc2["comments"][0]["n"]
    await call(s, "word_edit_comment", n=late_n, text="Add the distributor's name and start date.", **D)

    # Tracked change for the context view.
    await call(s, "word_track_changes", action="on", **D)
    await call(s, "word_replace", find="strongest", replace="best", **D)
    await call(s, "word_track_changes", action="off", **D)

    ctx = await raw(s, "word_context", "word_context (text)", **D)
    text = ctx.content[0].text
    print("\n" + "\n".join("    | " + line for line in text.splitlines()) + "\n")
    check("[p1 Title] # Sales Review" in text, "title rendered")
    check("## Wins" in text and "# Summary" in text, "heading levels")
    check("- New distributor in Spain" in text and "1. Hire a second sales rep" in text
          and "2. Launch the partner program" in text, "bullet and numbered lists")
    check("**every quarter**" in text, "bold kept")
    check("| Q4 | 2.4M |" in text, "table rendered")
    check("![image 1: Bar chart of quarterly sales, 8.0x4.0 cm]" in text, "image placeholder with alt text and size")
    check("{-*strongest*-}{+*best*+}" in text, "tracked change shown as {-del-}{+ins+} (italic kept)")
    check("[p4] ![image 1" in text, "image paragraph is plain, not a heading")
    lt = await call(s, "word_com", path=f'Documents("{doc}").Tables(1).Range.ListFormat.ListType')
    check(lt == 0, f"table cells are not list-numbered (ListType {lt})")
    check('reply [c' in text and "[resolved]" in text, "comment threads with replies and resolved flag")
    check("start date" in text, "edited comment text")
    check("paragraph numbers may drift" not in text, "paragraph numbering matches Word")

    # [pN] must be Word's paragraph N.
    rd = await call(s, "word_read", **D)
    words = {p["n"]: p["text"] for p in rd["paragraphs"]}
    import re
    bad = []
    for m in re.finditer(r"^\[p(\d+)[^\]]*\] (?:#+ |- |\d+\. )?(.+)$", text, re.M):
        n, shown = int(m.group(1)), re.sub(r"[*{}+\-]|\[c\d+\]", "", m.group(2)).strip()
        plain = re.sub(r"[*]", "", words.get(n, "")).strip()
        if "table" in shown[:6]:
            continue
        if shown[:12] and not plain.replace("strongestbest", "").startswith(shown[:12].replace("strongestbest", "")):
            if "image" not in shown:
                bad.append((n, shown[:30], plain[:30]))
    check(not bad, f"[pN] lines match word_read paragraph N {bad[:3]}")

    ctxi = await raw(s, "word_context", "word_context (with images)", include_images=True, **D)
    kinds = [c.type for c in ctxi.content]
    check("image" in kinds, f"context attaches images ({kinds})")

    gi = await raw(s, "word_get_image", n=1, **D)
    im = [c for c in gi.content if c.type == "image"]
    check(im and im[0].mimeType == "image/png", "word_get_image returns the PNG")
    meta = json.loads(gi.content[0].text)
    check(meta["pixels"] == [480, 240], f"original pixels kept {meta['pixels']}")

    li = await call(s, "word_list_images", **D)
    check(li["count"] == 1 and li["images"][0]["paragraph"] == 4, "word_list_images")

    pg = await raw(s, "word_page_image", page=1, **D)
    pim = [c for c in pg.content if c.type == "image"][0]
    with open(os.path.join(OUT, "context-test-page1.png"), "wb") as f:
        f.write(base64.b64decode(pim.data))
    check(json.loads(pg.content[0].text)["pixels"][0] > 700, "page image rendered")

    await call(s, "word_delete_comment", n=late_n, **D)
    await call(s, "word_delete_comment", n=0, **D)
    lc3 = await call(s, "word_list_comments", **D)
    check(lc3["comment_count"] == 0, "delete one, then delete all (works with replies)")
    await call(s, "word_delete_comment", "delete all on empty doc", n=0, **D)
    check(True, "delete all with no comments does not error")


asyncio.run(main())
