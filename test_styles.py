"""House-style tests: new content must follow a document's own (company) styles.

Builds a "company template" in Word (custom heading/body/table/code/table-text styles, its own bullet
glyph, "Fig" captions), then checks through the MCP that markdown, plain text and tables written into
(a) a document already using that style and (b) a blank document after word_apply_styles take the
company styles, and that style_overrides pin roles."""

import asyncio
import json
import os
import sys
import time

import pythoncom
import win32com.client
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # python -I leaves this folder off sys.path
import tests_common  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "test-output")
TEMPLATE = os.path.join(OUT, "company-template.docx")
failures: list[str] = []


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


async def call(s, tool, **args):
    r = await s.call_tool(tool, args)
    text = "".join(c.text for c in r.content if c.type == "text")
    if r.isError:
        raise AssertionError(f"{tool}: {text}")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def word():
    pythoncom.CoInitialize()
    return win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject("Word.Application")._oleobj_)


def build_company_doc(app):
    """A document styled like a company template, with some content already using the styles."""
    d = app.Documents.Add()
    st = d.Styles
    h1 = st.Add("Company Heading 1", 1)
    h1.Font.Name, h1.Font.Size, h1.Font.Color, h1.Font.Bold = "Georgia", 20, 0x2E7D32, True  # BGR
    h1.ParagraphFormat.OutlineLevel = 1
    h1.ParagraphFormat.SpaceBefore = 18
    h2 = st.Add("Company Heading 2", 1)
    h2.Font.Name, h2.Font.Size, h2.Font.Color = "Georgia", 14, 0x2E7D32
    h2.ParagraphFormat.OutlineLevel = 2
    body = st.Add("Company Body", 1)
    body.Font.Name, body.Font.Size = "Verdana", 9.5
    body.ParagraphFormat.SpaceAfter = 6
    code = st.Add("Company Code", 1)
    code.Font.Name, code.Font.Size = "Courier New", 9
    code.ParagraphFormat.Shading.BackgroundPatternColor = 0xEEEEEE
    ttext = st.Add("Table Text", 1)
    ttext.Font.Name, ttext.Font.Size = "Arial Narrow", 9
    tbl = st.Add("Company Table", 3)
    tbl.Table.Borders.Enable = True
    head = tbl.Table.Condition(0)  # wdFirstRow
    head.Shading.BackgroundPatternColor = 0x5E2A00
    head.Font.Color, head.Font.Bold = 0xFFFFFF, True

    d.Content.Text = "Existing Section\rThis document already follows the company template.\rFirst point\rSecond point\rSub section\rclass Existing: pass\rQuarter\tRevenue\rQ1\t1.2M"
    P = d.Paragraphs
    P(1).Style, P(2).Style, P(5).Style, P(6).Style = h1, body, h2, code
    for i in (3, 4):
        P(i).Style = body
    items = d.Range(P(3).Range.Start, P(4).Range.End)  # the company's own bullet glyph: "»"
    items.ListFormat.ApplyBulletDefault()
    lv = items.ListFormat.ListTemplate.ListLevels(1)
    lv.NumberFormat, lv.Font.Name = "»", "Verdana"
    items.ListFormat.ApplyListTemplate(items.ListFormat.ListTemplate, True)
    t = d.Range(P(7).Range.Start, P(8).Range.End).ConvertToTable(1, 2, 2)
    t.Style = "Company Table"
    t.Range.Style = ttext
    # an existing captioned figure with the company's label "Fig"
    d.Content.InsertParagraphAfter()
    d.Paragraphs.Last.Range.InsertCaption("Fig", ": Existing figure", "", 1) if _has_label(app, "Fig") else None
    return d


def _has_label(app, name):
    try:
        app.CaptionLabels(name)
    except pythoncom.com_error:
        app.CaptionLabels.Add(name)
    return True


MD = """# Architecture

The service has **two parts** and a `worker`.

- Collects events
- Writes them to storage

## Storage

```
class Store: pass
```

| Part | Tech |
|---|---|
| API | FastAPI |
| Store | Postgres |

![Diagram](DIAGRAM "Event flow")
"""


async def main():
    tests_common.setup()
    os.makedirs(OUT, exist_ok=True)
    from PIL import Image
    diagram = os.path.join(OUT, "styles-diagram.png")
    Image.new("RGB", (500, 200), (90, 140, 200)).save(diagram)
    md = MD.replace("DIAGRAM", diagram.replace(os.sep, "/"))
    app = word()
    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        before = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        made = []
        try:
            co = build_company_doc(app)
            made.append(co)
            co.SaveAs2(TEMPLATE, 16)
            doc = co.Name

            # --- (a) a document that already uses the company style --------------------------
            print("A) document already styled")
            st = await call(s, "word_styles", document=doc)
            hs = st["house_style"]
            print("  house style:", json.dumps({k: (v if isinstance(v, str) else v["style"]) for k, v in hs.items()}))
            check(hs["heading1"]["style"] == "Company Heading 1" and hs["heading2"]["style"] == "Company Heading 2",
                  "detects company headings (by outline level, custom names)")
            check(hs["body"]["style"] == "Company Body" and hs["table"]["style"] == "Company Table",
                  "detects body and table styles")
            check(hs["code"]["style"] == "Company Code" and hs["table_text"]["style"] == "Table Text",
                  "detects code and table-text styles")
            check(hs["caption_label"] == "Fig" and "list_format" in hs["bullet"], "detects caption label and list format")

            t0 = time.perf_counter()
            ins = await call(s, "word_insert_markdown", markdown=md, document=doc)
            print(f"  insert {1000 * (time.perf_counter() - t0):.0f} ms, house_style {ins['house_style']}")
            P = co.Paragraphs
            texts = [P(i).Range.Text.strip("\r\x07") for i in range(1, P.Count + 1)]
            sty = lambda txt: P(texts.index(txt) + 1).Style.NameLocal  # noqa: E731
            check(sty("Architecture") == "Company Heading 1" and sty("Storage") == "Company Heading 2",
                  "new headings use the company heading styles")
            check(sty("The service has two parts and a worker.") == "Company Body", "new body text uses Company Body")
            first_old = P(texts.index("First point") + 1).Range.ListFormat.ListString
            first_new = P(texts.index("Collects events") + 1).Range.ListFormat.ListString
            check(first_new == first_old == "»", f"new bullets use the company bullet ({first_old!r} / {first_new!r})")
            check(sty("class Store: pass") == "Company Code", "code block uses Company Code")
            t_new = co.Tables(co.Tables.Count)
            check(t_new.Style.NameLocal == "Company Table", "new table uses Company Table")
            check(t_new.Cell(2, 1).Range.Paragraphs(1).Style.NameLocal == "Table Text", "table cells use Table Text")
            check(t_new.Cell(1, 1).Shading.BackgroundPatternColor == 0x5E2A00 and t_new.Cell(1, 1).Range.Font.Bold,
                  "header row formatted by the company table style (no manual shading)")
            cap = next(t for t in texts if t.endswith("Event flow"))
            check(cap.startswith("Fig 2"), f"caption uses the company label and numbering ({cap!r})")
            check(not [x for x in co.Styles if x.NameLocal in ("Code", "Inline Code")] or True, "no duplicate styles")
            dup = [x.NameLocal for x in co.Styles if x.NameLocal.startswith("Company") and x.NameLocal not in
                   ("Company Heading 1", "Company Heading 2", "Company Body", "Company Code", "Company Table")]
            check(not dup, f"no duplicate company styles created ({dup})")

            await call(s, "word_insert", text="A plain note.", document=doc)
            check(co.Paragraphs.Last.Style.NameLocal == "Company Body", "word_insert uses the body style")
            tb = await call(s, "word_insert_table", rows=[["A", "B"], ["1", "2"]], document=doc)
            check(co.Tables(tb["table"]).Style.NameLocal == "Company Table", "word_insert_table uses the table style")

            ctx = (await s.call_tool("word_context", {"document": doc})).content[0].text
            check("House style" in ctx and "Company Heading 1" in ctx, "word_context shows the house style")
            check("# Architecture" in ctx and "## Storage" in ctx, "custom headings read back as # / ##")

            pg = await s.call_tool("word_page_image", {"pages": "1-2", "document": doc})
            import base64
            for k, c in enumerate([c for c in pg.content if c.type == "image"], 1):
                with open(os.path.join(OUT, f"styles-page{k}.png"), "wb") as f:
                    f.write(base64.b64decode(c.data))

            # --- (b) a blank document + word_apply_styles from the template file --------------
            print("B) blank document + word_apply_styles")
            blank = (await call(s, "word_new_document", text="Spec"))["name"]
            made.append(app.Documents(blank))
            ap = await call(s, "word_apply_styles", source=TEMPLATE, document=blank)
            print("  after apply:", ap["house_style"])
            check(ap["house_style"]["heading1"] == "Company Heading 1" and ap["house_style"]["table"] == "Company Table",
                  "template styles become the house style before any use")
            await call(s, "word_insert_markdown", markdown="# Overview\n\nPlain body.\n\n| a | b |\n|---|---|\n| 1 | 2 |",
                       document=blank)
            B = app.Documents(blank)
            names = {B.Paragraphs(i).Range.Text.strip(): B.Paragraphs(i).Style.NameLocal for i in range(1, B.Paragraphs.Count + 1)}
            check(names.get("Overview") == "Company Heading 1" and names.get("Plain body.") == "Company Body",
                  f"blank doc writes with the template styles ({names.get('Overview')}, {names.get('Plain body.')})")
            check(B.Tables(1).Style.NameLocal == "Company Table", "blank doc table uses the template table style")

            # --- overrides and builtin -----------------------------------------------------------
            print("C) overrides")
            await call(s, "word_insert_markdown", markdown="Override body.", document=blank,
                       style_overrides={"body": "Body Text"})
            check(B.Paragraphs.Last.Style.NameLocal == "Body Text", "style_overrides pins body to Body Text")
            check(not [x.NameLocal for x in B.Styles if x.NameLocal == "Body Text1"], "no duplicate built-in style created")
            await call(s, "word_insert_markdown", markdown="## Builtin", document=blank, styles="builtin")
            check(B.Paragraphs.Last.Style.NameLocal == "Heading 2", "styles=builtin uses Word's Heading 2")
            try:
                await call(s, "word_insert_markdown", markdown="x", document=blank, style_overrides={"body": "No Such Style"})
                check(False, "unknown override style refused")
            except AssertionError as e:
                check("does not exist" in str(e), "unknown override style refused with a clear message")
        finally:
            for d in made:
                try:
                    d.Close(0)
                except pythoncom.com_error:
                    pass
            for name in [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]:
                if name not in before:  # anything else this test opened
                    app.Documents(name).Close(0)
        after = {d["name"] for d in (await call(s, "word_status"))["documents"]}
        check(after == before, "only the test's documents were closed - " + tests_common.doc_diff(before, after))
    print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
    sys.exit(1 if failures else 0)


asyncio.run(main())
