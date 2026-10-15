"""Disposable fixture documents shaped like real integration-overview documents: section breaks,
layout and nested tables, an overview table and a references table repeating provider names,
a table of contents, headings, images with SEQ captions, comments, a content control, hidden text,
manual page breaks and a tracked deletion. Built through COM in the test process."""

import os

from PIL import Image, ImageDraw


def _png(path, label, color):
    im = Image.new("RGB", (640, 260), color)
    ImageDraw.Draw(im).text((20, 20), label, fill="white")
    im.save(path)
    return path


def build_integration_doc(app, out_dir: str):
    """Returns the open (unsaved) document."""
    os.makedirs(out_dir, exist_ok=True)
    img_vv = _png(os.path.join(out_dir, "fx-veloviewer.png"), "VELOVIEWER FLOW", (40, 90, 160))
    img_gsv = _png(os.path.join(out_dir, "fx-gsv.png"), "GSV FLOW", (40, 140, 90))
    d = app.Documents.Add()
    S = lambda sid: d.Styles(sid)  # noqa: E731
    rng = d.Content

    def para(text, style=-1):
        r = d.Content
        r.InsertParagraphAfter()
        p = d.Paragraphs(d.Paragraphs.Count).Range
        p.InsertAfter(text)
        p = d.Paragraphs(d.Paragraphs.Count)
        p.Style = S(style)
        p.Range.ListFormat.RemoveNumbers()
        return p

    def table(rows, style="Table Grid"):
        d.Content.InsertParagraphAfter()
        start = d.Paragraphs(d.Paragraphs.Count).Range.Start
        body = "\r".join("\t".join(r) for r in rows)
        d.Range(start, start).InsertAfter(body)
        d.Range(start, start + len(body)).Style = S(-1)
        t = d.Range(start, start + len(body)).ConvertToTable(1, len(rows), len(rows[0]))
        t.Style = style
        return t

    def image(path, caption):
        d.Content.InsertParagraphAfter()
        p = d.Paragraphs(d.Paragraphs.Count)
        p.Style = S(-1)
        shp = d.InlineShapes.AddPicture(path, False, True, p.Range)
        shp.Width = 300
        shp.Height = 120
        shp.Range.InsertCaption("Figure", ". " + caption, "", 1)

    rng.Text = "Integration Overview"
    d.Paragraphs(1).Style = S(-63)  # Title
    para("Contents", -1)
    para("References", -2)
    table([["Document", "Link"], ["D0180 - Integration Analysis - VeloViewer", "VeloViewer"],
           ["D0181 - Integration Analysis - Google Street View", "Google Street View"]])
    para("Overview", -2)
    t = table([["Provider", "Purpose", "Status"], ["VeloViewer", "Ride analytics", "Live"],
               ["Google Street View", "Imagery", "Planned"], ["Strava", "Activities", "Live"]])
    # a layout table with a nested table in one cell
    lt = table([["Layout left", "Layout right"]], style="Table Grid")
    inner = lt.Cell(1, 2).Range
    inner.Text = ""
    lt.Cell(1, 2).Tables.Add(lt.Cell(1, 2).Range, 2, 2)
    lt.Cell(1, 2).Tables(1).Cell(1, 1).Range.Text = "nested A"
    lt.Cell(1, 2).Tables(1).Cell(2, 2).Range.Text = "nested B"
    # next-page section break, landscape section
    d.Content.InsertParagraphAfter()
    end = d.Content.End - 1
    d.Range(end, end).InsertBreak(2)  # wdSectionBreakNextPage
    d.Sections(d.Sections.Count).PageSetup.Orientation = 1
    para("VeloViewer", -3)
    para("VeloViewer provides ride analytics. It is called from the API gateway.", -1)
    image(img_vv, "VeloViewer integration flow.")
    para("How it is called / request body", -4)
    para("POST /rides with a JSON body.", -1)
    para("Data received", -4)
    p = para("Ride segments and power data are returned.", -1)
    d.Comments.Add(p.Range, "Confirm the payload size.")
    # manual page break inside a paragraph (does not end it)
    p = para("Text before a page break", -1)
    p.Range.InsertAfter("")
    r = p.Range
    d.Range(r.End - 1, r.End - 1).InsertBreak(7)  # wdPageBreak
    d.Range(d.Paragraphs(d.Paragraphs.Count).Range.End - 1, d.Paragraphs(d.Paragraphs.Count).Range.End - 1).InsertAfter("text after the page break")
    # hidden paragraph
    h = para("Hidden note for editors.", -1)
    h.Range.Font.Hidden = True
    # back to portrait
    d.Content.InsertParagraphAfter()
    end = d.Content.End - 1
    d.Range(end, end).InsertBreak(2)
    d.Sections(d.Sections.Count).PageSetup.Orientation = 0
    para("Google Street View", -3)
    para("Google Street View supplies imagery for routes.", -1)
    image(img_gsv, "Google Street View integration flow.")
    p = para("Google Street View protocol", -4)
    d.Comments.Add(p.Range, "Which API version?")
    para("Requests are signed with a key.", -1)
    # content control holding two paragraphs
    para("Content control follows.", -1)
    d.Content.InsertParagraphAfter()
    last = d.Paragraphs(d.Paragraphs.Count).Range
    cc = d.ContentControls.Add(0, d.Range(last.Start, last.Start))  # wdContentControlRichText
    cc.Range.Text = "Inside control one\rInside control two"
    para("Conclusion", -2)
    para("Both providers are integrated through the gateway.", -1)
    # tracked deletion of a whole paragraph
    para("Paragraph to be deleted with tracking.", -1)
    d.TrackRevisions = True
    d.Paragraphs(d.Paragraphs.Count).Range.Delete()
    d.TrackRevisions = False
    para("End of document.", -1)
    # table of contents after "Contents"
    for i in range(1, d.Paragraphs.Count + 1):
        if d.Paragraphs(i).Range.Text.strip() == "Contents":
            r = d.Paragraphs(i).Range
            r.InsertParagraphAfter()
            tr = d.Paragraphs(i + 1).Range
            d.TablesOfContents.Add(d.Range(tr.Start, tr.Start), True, 1, 3)
            break
    d.Fields.Update()
    return d
