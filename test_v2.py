"""Tests for markdown writing, anchors/expect, word_changes (user edits made directly in Word),
images (data, svg, caption, replace), multi-page renders, TOC, batch, caching and big-document speed."""

import asyncio
import base64
import io
import json
import os
import re
import sys
import time

import pythoncom
import win32com.client
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


async def ctx_text(s, label="word_context", **args):
    r = await raw(s, "word_context", label, **args)
    return r.content[0].text


def make_png(path, w, h, label, color=(40, 90, 170)):
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    d.rectangle((10, 10, w - 10, h - 10), outline=color, width=6)
    d.text((30, 30), label, fill="black")
    im.save(path)


class UserWord:
    """The 'user': edits Word directly over COM, outside the MCP (like typing by hand)."""

    def __init__(self):
        pythoncom.CoInitialize()
        self.app = win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject("Word.Application")._oleobj_)

    def doc(self, name):
        return self.app.Documents(name)


async def main():
    os.makedirs(OUT, exist_ok=True)
    arch = os.path.join(OUT, "v2-arch.png")
    make_png(arch, 1400, 600, "ARCHITECTURE DIAGRAM")
    svg = os.path.join(OUT, "v2-flow.svg")
    with open(svg, "w") as f:
        f.write('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="120"><rect x="5" y="5" width="150" '
                'height="80" fill="#cde"/><text x="20" y="50">Ingest</text><rect x="240" y="5" width="150" '
                'height="80" fill="#edc"/><text x="260" y="50">Store</text></svg>')
    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        print(len((await s.list_tools()).tools), "tools")
        before = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        doc = (await call(s, "word_new_document", text="Platform Design"))["name"]
        user = UserWord()
        try:
            await scenario(s, doc, arch, svg, user)
            await big_document(s, user)
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


MD = """## Overview

The platform has **three tiers**: a *web client*, an `api` service and a database. Source: [repo](https://example.com/repo).

## Components

- Web client (React)
  - Talks to the API over HTTPS
- API (FastAPI)
- Database (Postgres)

## Request flow

1. Request arrives
2. Auth checks the token
   1. Validate signature
   2. Check expiry
3. Handler runs

```python
def handler(req):
    user = auth(req)
    return db.query(user)
```

| Layer | Tech | Notes |
|:--|:-:|--:|
| Client | React | SPA |
| API | FastAPI | `uvicorn` |

![System architecture](ARCH "System architecture"){width=14cm}

![Ingest flow](SVG "Ingest to store")

## Deployment

Runs on Kubernetes.
"""


async def scenario(s, doc, arch, svg, user):
    D = {"document": doc}
    st = await call(s, "word_status")
    me = next(x for x in st["documents"] if x["name"] == doc)
    check(me["cloud"] is False and "autosave" in me, f"status shows cloud/autosave ({me['cloud']}, {me['autosave']})")
    await call(s, "word_format", paragraph=1, style="Title", **D)

    md = MD.replace("ARCH", arch.replace(os.sep, "/")).replace("SVG", svg.replace(os.sep, "/"))
    ins = await call(s, "word_insert_markdown", "markdown: whole section", markdown=md, **D)
    print("  insert_markdown:", {k: ins[k] for k in ("first_paragraph", "last_paragraph", "headings", "tables", "images", "captions", "warnings", "ms")})
    check(ins["headings"] == 4 and ins["tables"] == 1 and ins["images"] == 2 and ins["captions"] == 2 and not ins["warnings"],
          "markdown: 4 headings, 1 table, 2 images, 2 captions, no warnings")

    text = await ctx_text(s, "word_context (first read)", **D)
    print("\n".join("    | " + l for l in text.splitlines()[:60]))
    check("## Components" in text and "- Web client (React)" in text and "  - Talks to the API over HTTPS" in text,
          "headings and nested bullets")
    check("1. Request arrives" in text and "  a. Validate signature" in text and "3. Handler runs" in text,
          "numbered list with lettered sub-list")
    check("[p" in text and "Code] def handler(req):" in text, "code block in Code style")
    check("| API | FastAPI | `uvicorn` |" in text or "| API | FastAPI | uvicorn |" in text, "table row")
    check("Figure 1: System architecture" in text and "Figure 2: Ingest to store" in text, "captions numbered 1 and 2")
    check(text.count("![image") == 2, "two images in context (PNG and SVG)")
    check("**three tiers**" in text and "[repo](https://example.com/repo)" in text, "bold and link kept")
    check("paragraph numbers may drift" not in text, "paragraph numbers verified")
    m = re.search(r"\[p(\d+)[^\]]*\] # Platform Design", text)
    check(m and m.group(1) == "1", "title stays paragraph 1")

    # Headings look like the document's own template (not the fragment fallback).
    look = await call(s, "word_com", path=f'Documents("{doc}").Paragraphs({ins["first_paragraph"]}).Range.Font.Name')
    check(look == "Calibri Light", f"heading uses the template font ({look})")

    t2 = time.perf_counter()
    await ctx_text(s, "word_context (cached)", **D)
    t_cached = (time.perf_counter() - t2) * 1000
    check(t_cached < 400, f"cached context read fast ({t_cached:.0f} ms)")

    # Anchors: section_end puts content at the end of "Components", before "Request flow".
    await call(s, "word_insert_markdown", "markdown: section_end", markdown="- Cache (Redis)", position="section_end",
               anchor="Components", **D)
    rd = await call(s, "word_read", **D)
    texts = [p["text"] for p in rd["paragraphs"]]
    i_cache, i_flow = texts.index("Cache (Redis)"), texts.index("Request flow")
    check(i_cache == i_flow - 1 and texts[i_cache - 1] == "Database (Postgres)", "section_end lands before next heading")

    # expect: refused when the paragraph does not match
    try:
        await call(s, "word_set_paragraph", paragraph=2, expect="Something else", text="x", **D)
        check(False, "expect mismatch refused")
    except AssertionError as e:
        check("document changed" in str(e), "expect mismatch refused with a clear message")
    await call(s, "word_set_paragraph", anchor="Runs on Kubernetes", text="Runs on Kubernetes (3 nodes).", **D)

    # Replace a whole section by heading.
    rs = await call(s, "word_replace_section", "replace section", anchor="Request flow",
                    markdown="The request goes **client -> API -> DB**.\n\n```\nGET /items\n```", **D)
    text = await ctx_text(s, "word_context (after replace)", **D)
    body = text.split("--- body")[1]
    check("client -> API -> DB" in body and "Request arrives" not in body and "## Deployment" in body
          and body.index("## Request flow") < body.index("client -> API") < body.index("## Deployment"),
          f"section replaced, heading and next section kept (removed {rs['removed_paragraphs']})")
    check("Figure" not in body.split("## Request flow")[1].split("## Deployment")[0],
          "images of the replaced section are gone with it")

    # word_changes: agent edits are not reported; user edits are.
    await ctx_text(s, "word_context (baseline)", **D)
    await call(s, "word_insert", text="Agent note.", anchor="Deployment", position="section_end", **D)
    ch = await call(s, "word_changes", "word_changes (after agent edit)", **D)
    check(not ch["changed"], f"agent's own edit not reported ({ch.get('changes')})")
    ud = user.doc(doc)
    p2 = ud.Paragraphs(2).Range
    ud.Range(p2.Start, p2.Start).InsertBefore("USER INTRO LINE\r")       # user types a new paragraph
    for i in range(1, ud.Paragraphs.Count + 1):
        if ud.Paragraphs(i).Range.Text.startswith("Web client"):
            r = ud.Paragraphs(i).Range
            ud.Range(r.End - 1, r.End - 1).InsertAfter(" - user edit")   # user changes a line
            break
    ud.InlineShapes.AddPicture(arch, False, True, ud.Paragraphs(ud.Paragraphs.Count).Range)
    ch = await call(s, "word_changes", "word_changes (user edits)", **D)
    print("  changes:", json.dumps(ch)[:600])
    kinds = {c["type"] for c in ch["changes"]}
    added = [c for c in ch["changes"] if c["type"] == "added" and "USER INTRO LINE" in c.get("now", [])]
    changed = [c for c in ch["changes"] if c["type"] == "changed" and any("user edit" in t for t in c.get("now", []))]
    check(ch["changed"] and added and added[0]["paragraphs"] == "2", "user's new paragraph reported as added at p2")
    check(bool(changed), f"user's edited line reported as changed ({kinds})")
    imgs = ch["counts"].get("images", {})
    check(imgs and imgs["now"] == imgs["before"] + 1, f"user's added image counted ({imgs})")
    ch2 = await call(s, "word_changes", "word_changes (nothing new)", **D)
    check(not ch2["changed"], "second word_changes is quiet")
    text = await ctx_text(s, "word_context (sees user edits)", **D)
    check("USER INTRO LINE" in text and "user edit" in text, "context includes the user's unsaved edits")

    # Images: base64 with caption, replace in place, svg by path.
    png_b = io.BytesIO()
    Image.new("RGB", (600, 300), (200, 60, 60)).save(png_b, "PNG")
    b64 = base64.b64encode(png_b.getvalue()).decode()
    im1 = await call(s, "word_insert_image", "insert image (base64 + caption)", data_base64=b64, position="section_end",
                     anchor="Overview", caption="Red block", alt_text="red block", **D)
    check(im1["caption"] and im1["width_cm"] > 0, f"base64 image with caption ({im1})")
    li = await call(s, "word_list_images", **D)
    red = next(m for m in li["images"] if m["alt_text"] == "red block")
    blue = os.path.join(OUT, "v2-blue.png")
    make_png(blue, 300, 300, "BLUE")
    rep = await call(s, "word_insert_image", "replace image", path=blue, replace=red["n"], **D)
    li2 = await call(s, "word_list_images", **D)
    # a square picture replacing a 2:1 one fits the old box: height-limited, so width = old height
    check(li2["count"] == li["count"] and abs(rep["height_cm"] - red["height_cm"]) < 0.05
          and abs(rep["width_cm"] - red["height_cm"]) < 0.05,
          f"image replaced in place, fitted to the old box ({red['width_cm']}x{red['height_cm']} -> "
          f"{rep['width_cm']}x{rep['height_cm']})")
    gi = await raw(s, "word_get_image", n=red["n"], **D)
    px = json.loads(gi.content[0].text)["pixels"]
    check(px == [300, 300], f"replaced image is the new picture ({px})")
    text = await ctx_text(s, "word_context", **D)
    check("Figure 1: Red block" in text or "Figure 1: System architecture" in text, "captions present")

    # TOC + fields
    await call(s, "word_insert_toc", position="after", anchor="Platform Design", **D)
    # a new captioned figure ahead of "Red block": numbers are stale until update_fields
    await call(s, "word_insert_markdown", markdown=f"![Earlier]({arch.replace(os.sep, '/')} \"Earlier figure\"){{width=6cm}}",
               position="after", anchor="Overview", **D)
    upd = await call(s, "word_update_fields", **D)
    check(upd["tables_of_contents"] == 1, "table of contents inserted and updated")
    text = await ctx_text(s, "word_context (with TOC)", **D)
    check("Figure 1: Earlier figure" in text and "Figure 2: Red block" in text,
          "caption numbers renumbered after update_fields")

    # Multi-page render
    pg = await raw(s, "word_page_image", "page images 1-2", pages="1-2", **D)
    imgs = [c for c in pg.content if c.type == "image"]
    check(len(imgs) == 2, f"two page images ({len(imgs)})")
    for k, c in enumerate(imgs, 1):
        with open(os.path.join(OUT, f"v2-page{k}.png"), "wb") as f:
            f.write(base64.b64decode(c.data))

    # Batch
    b = await call(s, "word_batch", "batch of 3", calls=[
        {"tool": "word_insert_markdown", "args": {"markdown": "## Risks\n\n- Single region", "document": doc}},
        {"tool": "word_format", "args": {"find": "Single region", "bold": True, "document": doc}},
        {"tool": "word_list_images", "args": {"document": doc}},
    ])
    check(b["ran"] == 3 and all(x["ok"] for x in b["results"]), "batch ran 3 calls")
    bad = await call(s, "word_batch", calls=[{"tool": "word_set_paragraph", "args": {"anchor": "nope", "text": "x", "document": doc}},
                                             {"tool": "word_list_images", "args": {"document": doc}}])
    check(bad["ran"] == 1 and not bad["results"][0]["ok"], "batch stops at first error")

    # Undo removes one whole markdown insert
    n_before = (await call(s, "word_read", **D))["paragraph_count"]
    await call(s, "word_insert_markdown", markdown="## Temp\n\n- a\n- b\n\n| x | y |\n|---|---|\n| 1 | 2 |", **D)
    await call(s, "word_undo", **D)
    n_after = (await call(s, "word_read", **D))["paragraph_count"]
    check(n_after == n_before, f"one undo removes a whole markdown insert ({n_before} -> {n_after})")


async def big_document(s, user):
    """~40 screenshots + 300 paragraphs: how fast are reads and writes on a heavy document?"""
    print("\n  big document:")
    shots = []
    for i in range(8):
        p = os.path.join(OUT, f"v2-shot{i}.png")
        im = Image.new("RGB", (1600, 1000), (240, 240, 245))
        dr = ImageDraw.Draw(im)
        for k in range(0, 1000, 40):
            dr.line((0, k, 1600, k + (i * 7) % 40), fill=(180 - i * 10, 190, 200), width=2)
        dr.text((50, 50), f"SCREEN {i}", fill="black")
        im.save(p)
        shots.append(p.replace(os.sep, "/"))
    md = []
    for sec in range(40):
        md.append(f"## Module {sec}\n\nThe module {sec} handles part of the request pipeline. " * 1)
        md.append("\n".join(f"- Responsibility {k} of module {sec}" for k in range(5)))
        md.append(f"![Screen {sec}]({shots[sec % 8]} \"Module {sec} screen\"){{width=15cm}}")
    doc = (await call(s, "word_new_document", text="Big Spec"))["name"]
    D = {"document": doc}
    t = time.perf_counter()
    ins = await call(s, "word_insert_markdown", "BIG insert 40 sections + 40 screenshots", markdown="\n\n".join(md), **D)
    t_ins = time.perf_counter() - t
    print(f"    insert: {t_ins:.2f}s, {ins['paragraph_count']} paragraphs, {ins['images']} images, server {ins['ms']} ms")
    check(ins["images"] == 40, "40 screenshots inserted")
    for label, args in (("BIG context first (stats)", {"refresh": True}), ("BIG context cached", {}),
                        ("BIG context no stats", {"stats": False, "refresh": True})):
        t = time.perf_counter()
        await ctx_text(s, label, **args, **D)
        print(f"    {label}: {(time.perf_counter() - t) * 1000:.0f} ms")
    t = time.perf_counter()
    await call(s, "word_changes", "BIG word_changes", **D)
    print(f"    word_changes: {(time.perf_counter() - t) * 1000:.0f} ms")
    t = time.perf_counter()
    await call(s, "word_insert_markdown", "BIG insert one section", markdown="## Extra\n\nOne more section.", anchor="Module 20",
               position="section_end", **D)
    print(f"    insert one section mid-document: {(time.perf_counter() - t) * 1000:.0f} ms")


asyncio.run(main())
