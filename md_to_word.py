"""Markdown -> Word flat-OPC fragment for Range.InsertXML (one COM call for a whole section).

Supports headings, paragraphs, bold/italic/strike, inline code, links, bullet/numbered lists
(nested, each list restarts its numbering), block quotes, fenced code blocks, pipe tables,
horizontal rules and images (local path, data: URI or http(s) URL; optional
`{width=8cm}` / `{width=50%}` after the image; the image title becomes a numbered caption).
SVG images cannot be embedded through XML without a raster fallback, so they are returned as
placeholders for the caller to insert with InlineShapes.AddPicture.
"""

from __future__ import annotations

import base64
import io
import os
import re
import urllib.request
from dataclasses import dataclass, field
from xml.sax.saxutils import escape, quoteattr

from markdown_it import MarkdownIt

EMU_PER_CM = 360000
MAX_DOWNLOAD = 25 * 1024 * 1024
SVG_MARK = "⁣WMCPSVG{}⁣"  # invisible separator brackets: never typed by people

NS_DECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)

# Styles the fragment may reference. Word keeps the destination's definition when a style
# with the same name already exists, so these only fill in what the document lacks.
STYLES = """<w:styles %s>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="40"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="40"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="4"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="5"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:i/><w:color w:val="44546A"/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepLines/><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="144" w:right="144"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:noProof/><w:sz w:val="19"/></w:rPr></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="InlineCode"><w:name w:val="Inline Code"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:noProof/><w:color w:val="C7254E"/><w:sz w:val="19"/><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>""" % NS_DECL

BULLETS = [("", "Symbol"), ("o", "Courier New"), ("", "Wingdings")]
NUMFMT = [("decimal", "%{}."), ("lowerLetter", "%{}."), ("lowerRoman", "%{}.")]


def _numbering(lists: list[tuple[int, str]]) -> str:
    """One abstractNum per list kind, one num per list so every list restarts at 1."""
    lv_b, lv_n = [], []
    for i in range(9):
        ch, font = BULLETS[i % 3]
        ind = f'<w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="360"/></w:pPr>'
        lv_b.append(f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                    f'<w:lvlText w:val="{ch}"/><w:lvlJc w:val="left"/>{ind}'
                    f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr></w:lvl>')
        fmt, txt = NUMFMT[i % 3]
        lv_n.append(f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                    f'<w:lvlText w:val="{txt.format(i + 1)}"/><w:lvlJc w:val="left"/>{ind}</w:lvl>')
    out = [f'<w:numbering {NS_DECL}>',
           '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' + "".join(lv_b) + "</w:abstractNum>",
           '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>' + "".join(lv_n) + "</w:abstractNum>"]
    for num_id, kind in lists:
        absid = 1 if kind == "bullet" else 2
        out.append(f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{absid}"/>'
                   + "".join(f'<w:lvlOverride w:ilvl="{i}"><w:startOverride w:val="1"/></w:lvlOverride>' for i in range(9))
                   + "</w:num>")
    out.append("</w:numbering>")
    return "".join(out)


@dataclass
class Fragment:
    xml: str
    paragraphs: int
    headings: int = 0
    tables: int = 0
    images: int = 0
    svg: list = field(default_factory=list)       # [(marker, path, width_cm, alt)]
    captions: int = 0
    warnings: list = field(default_factory=list)


class Builder:
    def __init__(self, base_dir: str, max_width_cm: float):
        self.base_dir = base_dir
        self.max_width_cm = max_width_cm
        self.body: list[str] = []
        self.rels: list[str] = []
        self.media: list[tuple[str, str, bytes]] = []
        self.lists: list[tuple[int, str]] = []
        self.paras = 0
        self.headings = 0
        self.tables = 0
        self.images = 0
        self.captions = 0
        self.svg: list = []
        self.warnings: list[str] = []
        self.widths: list = []
        self._rid = 0
        self._docpr = 9000

    def rid(self) -> str:
        self._rid += 1
        return f"rIdM{self._rid}"

    # --- images --------------------------------------------------------------------
    def load_image(self, src: str) -> tuple[bytes, str]:
        if src.startswith("data:"):
            head, _, data = src.partition(",")
            return base64.b64decode(data), head[5:].split(";")[0]
        if re.match(r"https?://", src, re.I):
            req = urllib.request.Request(src, headers={"User-Agent": "word-mcp"})
            with urllib.request.urlopen(req, timeout=20) as r:
                ctype = r.headers.get_content_type()
                data = r.read(MAX_DOWNLOAD + 1)
            if len(data) > MAX_DOWNLOAD:
                raise ValueError(f"image larger than {MAX_DOWNLOAD // 2**20} MB: {src}")
            if not ctype.startswith("image/"):
                raise ValueError(f"not an image ({ctype}): {src}")
            return data, ctype
        path = src
        if path.startswith("file:///"):
            path = urllib.request.url2pathname(path[8:])
        if not os.path.isabs(path):
            path = os.path.join(self.base_dir, path)
        with open(path, "rb") as fh:
            data = fh.read()
        ext = os.path.splitext(path)[1].lower()
        return data, {".svg": "image/svg+xml", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                      ".gif": "image/gif", ".emf": "image/x-emf", ".wmf": "image/x-wmf"}.get(ext, "")

    def image_run(self, src: str, alt: str) -> str:
        width_hint = self.widths.pop(0) if self.widths else None
        try:
            data, ctype = self.load_image(src)
        except Exception as e:  # noqa: BLE001 - keep writing, report it
            self.warnings.append(f"image {src!r} skipped: {e}")
            return self.text_run(f"[missing image: {alt or src}]", {})
        self.images += 1
        if ctype == "image/svg+xml" or data[:200].lstrip().startswith((b"<svg", b"<?xml")):
            marker = SVG_MARK.format(len(self.svg) + 1)
            path = src if not re.match(r"(https?|data):", src) else None
            if path is None:  # remote/data SVG: write a temp file for AddPicture
                import tempfile
                fd, path = tempfile.mkstemp(suffix=".svg")
                os.write(fd, data)
                os.close(fd)
            elif not os.path.isabs(path):
                path = os.path.join(self.base_dir, path)
            self.svg.append((marker, path, self._width_cm(width_hint, None), alt))
            return f'<w:r><w:t xml:space="preserve">{marker}</w:t></w:r>'
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        if im.format not in ("PNG", "JPEG", "GIF"):
            if im.format in ("WMF", "EMF"):
                im.load(dpi=144)
            buf = io.BytesIO()
            im.convert("RGBA").save(buf, "PNG")
            data, ctype = buf.getvalue(), "image/png"
            im = Image.open(io.BytesIO(data))
        ext = {"PNG": "png", "JPEG": "jpeg", "GIF": "gif"}[im.format]
        ctype = f"image/{ext}"
        dpi = (im.info.get("dpi") or (96, 96))[0] or 96
        nat_w = im.size[0] / float(dpi) * 2.54
        nat_h = im.size[1] / float(dpi) * 2.54
        w_cm = self._width_cm(width_hint, nat_w)
        h_cm = nat_h * w_cm / nat_w
        rid = self.rid()
        name = f"/word/media/wmcp{len(self.media) + 1}.{ext}"
        self.media.append((name, ctype, data))
        self.rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                         f'relationships/image" Target="media/{name.rsplit("/", 1)[1]}"/>')
        cx, cy = int(w_cm * EMU_PER_CM), int(h_cm * EMU_PER_CM)
        self._docpr += 1
        a = quoteattr(alt or "")
        return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                f'<wp:docPr id="{self._docpr}" name="Picture {self._docpr}" descr={a}/>'
                '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.{ext}" descr={a}/><pic:cNvPicPr/></pic:nvPicPr>'
                f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

    def _width_cm(self, hint, natural):
        if hint:
            val, unit = hint
            return min(self.max_width_cm, val if unit == "cm" else self.max_width_cm * val / 100.0)
        if natural is None:
            return self.max_width_cm * 0.8
        return min(natural, self.max_width_cm)

    # --- runs ----------------------------------------------------------------------
    def text_run(self, text: str, fmt: dict) -> str:
        props = []
        if fmt.get("code"):
            props.append('<w:rStyle w:val="InlineCode"/>')
        elif fmt.get("link"):
            props.append('<w:rStyle w:val="Hyperlink"/>')
        if fmt.get("b"):
            props.append("<w:b/>")
        if fmt.get("i"):
            props.append("<w:i/>")
        if fmt.get("s"):
            props.append("<w:strike/>")
        rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
        parts = text.split("\t")
        body = '<w:tab/>'.join(f'<w:t xml:space="preserve">{escape(p)}</w:t>' for p in parts)
        return f"<w:r>{rpr}{body}</w:r>"

    def inline(self, tok) -> str:
        if tok is None or not tok.children:
            return ""
        out, fmt, link_runs, link_href = [], {}, None, None
        for c in tok.children:
            t = c.type
            target = link_runs if link_runs is not None else out
            if t == "text":
                target.append(self.text_run(c.content, fmt))
            elif t == "code_inline":
                target.append(self.text_run(c.content, {**fmt, "code": True}))
            elif t == "softbreak":
                target.append(self.text_run(" ", fmt))
            elif t == "hardbreak":
                target.append("<w:r><w:br/></w:r>")
            elif t in ("strong_open", "strong_close"):
                fmt["b"] = t == "strong_open"
            elif t in ("em_open", "em_close"):
                fmt["i"] = t == "em_open"
            elif t in ("s_open", "s_close"):
                fmt["s"] = t == "s_open"
            elif t == "link_open":
                link_href = c.attrGet("href") or ""
                link_runs = []
                fmt["link"] = True
            elif t == "link_close":
                fmt["link"] = False
                runs = "".join(link_runs or [])
                if link_href.startswith("#"):
                    out.append(f'<w:hyperlink w:anchor={quoteattr(link_href[1:])}>{runs}</w:hyperlink>')
                elif link_href:
                    rid = self.rid()
                    self.rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/'
                                     f'2006/relationships/hyperlink" Target={quoteattr(link_href)} TargetMode="External"/>')
                    out.append(f'<w:hyperlink r:id="{rid}" w:history="1">{runs}</w:hyperlink>')
                else:
                    out.append(runs)
                link_runs, link_href = None, None
            elif t == "image":
                alt = "".join(x.content for x in (c.children or []) if x.type == "text") or c.content
                target.append(self.image_run(c.attrGet("src") or "", alt))
            elif t in ("html_inline",):
                target.append(self.text_run(c.content, fmt))
        return "".join(out)

    # --- blocks --------------------------------------------------------------------
    def para(self, runs: str, style: str = "", extra_ppr: str = "") -> None:
        ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + extra_ppr
        self.body.append(f"<w:p>{f'<w:pPr>{ppr}</w:pPr>' if ppr else ''}{runs}</w:p>")
        self.paras += 1

    def caption(self, title: str) -> None:
        runs = (self.text_run("Figure ", {}) +
                '<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
                + self.text_run(": " + title, {}))
        self.para(runs, "Caption", '<w:jc w:val="center"/>')  # under the centered image
        self.captions += 1

    def build(self, tokens) -> None:
        i, n = 0, len(tokens)
        list_stack: list[dict] = []   # {"kind", "num"}
        item_first: list[bool] = []
        quote = 0
        while i < n:
            t = tokens[i]
            ty = t.type
            if ty == "heading_open":
                level = int(t.tag[1])
                self.para(self.inline(tokens[i + 1]), f"Heading{level}")
                self.headings += 1
                i += 3
                continue
            if ty in ("bullet_list_open", "ordered_list_open"):
                kind = "bullet" if ty == "bullet_list_open" else "number"
                if list_stack and list_stack[0]["kind"] == kind:
                    num = list_stack[0]["num"]  # deeper level of the same list (Word restarts it per parent item)
                else:
                    num = len(self.lists) + 1   # new list: its own numbering, starting at 1
                    self.lists.append((num, kind))
                list_stack.append({"kind": kind, "num": num})
                i += 1
                continue
            if ty in ("bullet_list_close", "ordered_list_close"):
                list_stack.pop()
                i += 1
                continue
            if ty == "list_item_open":
                item_first.append(True)
                i += 1
                continue
            if ty == "list_item_close":
                item_first.pop()
                i += 1
                continue
            if ty == "blockquote_open":
                quote += 1
                i += 1
                continue
            if ty == "blockquote_close":
                quote -= 1
                i += 1
                continue
            if ty == "paragraph_open":
                inl = tokens[i + 1]
                kids = [c for c in (inl.children or []) if not (c.type == "text" and not c.content.strip())]
                lone_image = len(kids) == 1 and kids[0].type == "image" and not list_stack
                runs = self.inline(inl)
                if list_stack:
                    lvl = len(list_stack) - 1
                    if item_first and item_first[-1]:
                        num = list_stack[-1]["num"]
                        self.para(runs, "ListParagraph",
                                  f'<w:numPr><w:ilvl w:val="{lvl}"/><w:numId w:val="{num}"/></w:numPr>')
                        item_first[-1] = False
                    else:
                        self.para(runs, "ListParagraph", f'<w:ind w:left="{720 * (lvl + 1)}"/>')
                elif quote:
                    self.para(runs, "Quote")
                elif lone_image:
                    self.para(runs, "", '<w:keepNext/><w:jc w:val="center"/>' if kids[0].attrGet("title") else '<w:jc w:val="center"/>')
                    title = kids[0].attrGet("title")
                    if title:
                        self.caption(title)
                else:
                    self.para(runs)
                i += 3
                continue
            if ty in ("fence", "code_block"):
                lines = t.content.rstrip("\n").split("\n")
                for k, line in enumerate(lines):
                    keep = '<w:keepNext/>' if k < len(lines) - 1 else '<w:spacing w:after="160"/>'
                    self.para(self.text_run(line, {}) if line else "", "Code", keep)
                i += 1
                continue
            if ty == "table_open":
                i = self.table(tokens, i)
                continue
            if ty == "hr":
                self.para("", "", '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A6A6A6"/></w:pBdr>')
                i += 1
                continue
            if ty == "html_block":
                for line in t.content.rstrip("\n").split("\n"):
                    self.para(self.text_run(line, {}))
                i += 1
                continue
            i += 1

    def table(self, tokens, i) -> int:
        rows, aligns, cur, header_rows = [], [], None, 0
        in_head = False
        while tokens[i].type != "table_close":
            t = tokens[i]
            if t.type == "thead_open":
                in_head = True
            elif t.type == "thead_close":
                in_head = False
            elif t.type == "tr_open":
                cur = []
            elif t.type == "tr_close":
                rows.append(cur)
                header_rows += in_head
            elif t.type in ("th_open", "td_open"):
                style = t.attrGet("style") or ""
                m = re.search(r"text-align:(\w+)", style)
                cur.append({"runs": self.inline(tokens[i + 1]), "head": t.type == "th_open",
                            "align": {"center": "center", "right": "right"}.get(m.group(1) if m else "", "")})
            i += 1
        cols = max(len(r) for r in rows) if rows else 0
        xml = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>'
               '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
               '</w:tblPr><w:tblGrid>' + "".join('<w:gridCol/>' for _ in range(cols)) + "</w:tblGrid>"]
        for r_i, row in enumerate(rows):
            trpr = "<w:trPr><w:tblHeader/></w:trPr>" if r_i < header_rows else ""
            xml.append(f"<w:tr>{trpr}")
            for c_i in range(cols):
                cell = row[c_i] if c_i < len(row) else {"runs": "", "head": False, "align": ""}
                runs = cell["runs"]
                if cell["head"]:
                    runs = runs.replace("<w:r><w:rPr>", "<w:r><w:rPr><w:b/>").replace("<w:r><w:t", "<w:r><w:rPr><w:b/></w:rPr><w:t")
                shade = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if cell["head"] else ""
                jc = f'<w:jc w:val="{cell["align"]}"/>' if cell["align"] else ""
                xml.append(f'<w:tc><w:tcPr>{shade}</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>{runs}</w:p></w:tc>')
                self.paras += 1
            xml.append("</w:tr>")
            self.paras += 1  # end-of-row mark
        xml.append("</w:tbl>")
        self.body.append("".join(xml))
        self.tables += 1
        return i + 1

    # --- package ------------------------------------------------------------------
    def package(self) -> str:
        rels = ['<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
        if self.lists:
            rels.append('<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>')
        rels += self.rels
        parts = [
            '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships></pkg:xmlData></pkg:part>',
            '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">'
            '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels) + "</Relationships></pkg:xmlData></pkg:part>",
            '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
            f'<pkg:xmlData><w:document {NS_DECL}><w:body>' + "".join(self.body) + "</w:body></w:document></pkg:xmlData></pkg:part>",
            '<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">'
            f"<pkg:xmlData>{STYLES}</pkg:xmlData></pkg:part>",
        ]
        if self.lists:
            parts.append('<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">'
                         f"<pkg:xmlData>{_numbering(self.lists)}</pkg:xmlData></pkg:part>")
        for name, ctype, data in self.media:
            parts.append(f'<pkg:part pkg:name="{name}" pkg:contentType="{ctype}" pkg:compression="store">'
                         f"<pkg:binaryData>{base64.b64encode(data).decode()}</pkg:binaryData></pkg:part>")
        return ('<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?>'
                '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' + "".join(parts) + "</pkg:package>")


_WIDTH = re.compile(r"(!\[[^\]]*\]\([^)]*\))\{\s*width\s*=\s*([0-9.]+)\s*(cm|%)\s*\}")


def build(markdown: str, base_dir: str = ".", max_width_cm: float = 16.0) -> Fragment:
    """Markdown -> flat-OPC XML for Range.InsertXML."""
    widths: list = []
    # one slot per image (None = no hint) so each {width=...} lines up with its own image
    md_text = markdown.replace("\r\n", "\n")
    out, pos = [], 0
    for m in re.finditer(r"!\[[^\]]*\]\([^)]*\)(\{\s*width\s*=\s*[0-9.]+\s*(?:cm|%)\s*\})?", md_text):
        out.append(md_text[pos:m.start()])
        if m.group(1):
            w = _WIDTH.match(m.group(0))
            widths.append((float(w.group(2)), w.group(3)))
            out.append(w.group(1))
        else:
            widths.append(None)
            out.append(m.group(0))
        pos = m.end()
    out.append(md_text[pos:])
    md_text = "".join(out)

    md = MarkdownIt("commonmark", {"html": False}).enable("table").enable("strikethrough")
    b = Builder(base_dir, max_width_cm)
    b.widths = widths
    b.build(md.parse(md_text))
    return Fragment(b.package(), b.paras, b.headings, b.tables, b.images, b.svg, b.captions, b.warnings)
