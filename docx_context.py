"""Turn Word's flat-OPC package (Document.WordOpenXML, one COM call) into agent context:
markdown-like text with Word paragraph numbers, tables, images, comment threads, tracked
changes, footnotes, headers/footers and properties, plus the raw image bytes."""

from __future__ import annotations

import base64
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

NS = {
    "pkg": "http://schemas.microsoft.com/office/2006/xmlPackage",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "dc": "http://purl.org/dc/elements/1.1/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dcterms": "http://purl.org/dc/terms/",
}
EMU_PER_CM = 360000
for _p, _u in NS.items():  # serialize copied XML (list definitions) with the usual prefixes
    if _p not in ("rel",):
        ET.register_namespace(_p, _u)


def q(tag: str) -> str:
    p, local = tag.split(":")
    return "{%s}%s" % (NS[p], local)


W = lambda local: "{%s}%s" % (NS["w"], local)  # noqa: E731


def on(el, name: str) -> bool:
    """Boolean run/paragraph property (<w:b/>, <w:b w:val="0"/>)."""
    x = el.find(W(name)) if el is not None else None
    return x is not None and x.get(W("val"), "true") not in ("0", "false", "off")


@dataclass
class Piece:
    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    change: str = ""   # "ins" / "del"
    link: str = ""
    raw: bool = False  # tokens like images: never wrapped in markers


@dataclass
class Package:
    parts: dict = field(default_factory=dict)      # name -> Element (xml parts)
    binary: dict = field(default_factory=dict)     # name -> (content_type, base64 text)

    @classmethod
    def parse(cls, flat_xml: str) -> "Package":
        flat_xml = re.sub(r"^\s*<\?xml[^>]*\?>", "", flat_xml)
        root = ET.fromstring(flat_xml)
        pkg = cls()
        for part in root.findall(q("pkg:part")):
            name = part.get(q("pkg:name"))
            xd = part.find(q("pkg:xmlData"))
            if xd is not None and len(xd):
                pkg.parts[name] = xd[0]
                continue
            bd = part.find(q("pkg:binaryData"))
            if bd is not None and bd.text:
                pkg.binary[name] = (part.get(q("pkg:contentType"), ""), bd.text)  # base64, decoded on demand
        return pkg

    def rels(self, part: str) -> dict:
        folder, _, fname = part.rpartition("/")
        rp = self.parts.get(f"{folder}/_rels/{fname}.rels")
        out = {}
        if rp is None:
            return out
        for r in rp.findall(q("rel:Relationship")):
            target = r.get("Target", "")
            if r.get("TargetMode") != "External" and not target.startswith("/"):
                target = f"{folder}/{target}"
            out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], target)
        return out


class Context:
    def __init__(self, flat_xml: str):
        self.pkg = Package.parse(flat_xml)
        self.doc = self.pkg.parts.get("/word/document.xml")
        if self.doc is None:
            raise ValueError("no /word/document.xml in package")
        self.rels = self.pkg.rels("/word/document.xml")
        self._load_styles()
        self._load_numbering()
        self.images: list[dict] = []
        self.comment_order: list[str] = []          # XML comment ids in reference order
        self.comment_anchor: dict[str, list[str]] = {}
        self.comment_para: dict[str, int] = {}
        self.open_comments: set[str] = set()
        self.field_stack: list[str] = []
        self.footnote_refs: list[str] = []
        self.counters: dict = {}
        self.changes = {"ins": 0, "del": 0}
        self.lines: list[tuple[int, str]] = []      # (paragraph number, rendered line)
        self.heading_level: dict[int, int] = {}     # paragraph number -> 1..9 (Title = 0)
        self.inline_images = 0
        self.para = 0
        self.tables = 0
        self.para_meta: list[dict] = []             # per paragraph: style id, role facts (house style)
        self.table_styles: list[str] = []           # style id of each top-level table
        self.char_styles: dict[str, int] = {}       # character style id -> uses
        self._render_body()
        self.seq_labels: dict[str, int] = {}        # caption labels in use (SEQ Figure, SEQ Table...)
        for el in self.doc.iter():
            instr = el.get(W("instr")) if el.tag == W("fldSimple") else (el.text if el.tag == W("instrText") else None)
            m = re.search(r"SEQ\s+(\S+)", instr or "")
            if m:
                self.seq_labels[m.group(1)] = self.seq_labels.get(m.group(1), 0) + 1

    # --- styles / numbering ----------------------------------------------------
    def _load_styles(self):
        self.style_name, self.style_based, self.style_num = {}, {}, {}
        self.style_type, self.style_custom, self.style_outline, self.style_look = {}, set(), {}, {}
        st = self.pkg.parts.get("/word/styles.xml")
        if st is None:
            return
        for s in st.findall(W("style")):
            sid = s.get(W("styleId"))
            n = s.find(W("name"))
            self.style_name[sid] = n.get(W("val")) if n is not None else sid
            self.style_type[sid] = s.get(W("type"), "paragraph")
            if s.get(W("customStyle")) in ("1", "true"):
                self.style_custom.add(sid)
            b = s.find(W("basedOn"))
            if b is not None:
                self.style_based[sid] = b.get(W("val"))
            num = s.find(f"{W('pPr')}/{W('numPr')}")
            if num is not None:
                self.style_num[sid] = num
            ol = s.find(f"{W('pPr')}/{W('outlineLvl')}")
            if ol is not None:
                self.style_outline[sid] = int(ol.get(W("val"), "9"))
            rpr = s.find(W("rPr"))
            if rpr is not None:
                look = {}
                f = rpr.find(W("rFonts"))
                if f is not None and (f.get(W("ascii")) or f.get(W("asciiTheme"))):
                    look["font"] = f.get(W("ascii")) or f.get(W("asciiTheme"))
                sz = rpr.find(W("sz"))
                if sz is not None:
                    look["size_pt"] = int(sz.get(W("val"), "0")) / 2
                c = rpr.find(W("color"))
                if c is not None and c.get(W("val")) not in (None, "auto"):
                    look["color"] = "#" + c.get(W("val"))
                if on(rpr, "b"):
                    look["bold"] = True
                if look:
                    self.style_look[sid] = look

    def heading_level_of(self, sid) -> int | None:
        """1-9 for heading styles: built-in "heading N" or any style with an outline level
        (a company's own heading styles), following basedOn."""
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            m = re.match(r"heading (\d)$", self.style_name.get(sid, "").lower())
            if m:
                return int(m.group(1))
            if sid in self.style_outline and self.style_outline[sid] < 9:
                return self.style_outline[sid] + 1
            sid = self.style_based.get(sid)
        return None

    def _num_kind(self, numpr) -> str | None:
        if numpr is None:
            return None
        nid = numpr.find(W("numId"))
        lv = numpr.find(W("ilvl"))
        num_id = nid.get(W("val")) if nid is not None else None
        if num_id in (None, "0"):
            return None
        fmt = self.abs_fmt.get((self.num_abs.get(num_id), lv.get(W("val")) if lv is not None else "0"), "bullet")
        return "bullet" if fmt in ("bullet", "none") else "number"

    def _style_numpr(self, sid):
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            if sid in self.style_num:
                return self.style_num[sid]
            sid = self.style_based.get(sid)
        return None

    def _load_numbering(self):
        self.num_abs, self.abs_fmt = {}, {}
        nm = self.pkg.parts.get("/word/numbering.xml")
        if nm is None:
            return
        for a in nm.findall(W("abstractNum")):
            aid = a.get(W("abstractNumId"))
            for lvl in a.findall(W("lvl")):
                f = lvl.find(W("numFmt"))
                self.abs_fmt[(aid, lvl.get(W("ilvl")))] = f.get(W("val")) if f is not None else "decimal"
        for n in nm.findall(W("num")):
            a = n.find(W("abstractNumId"))
            if a is not None:
                self.num_abs[n.get(W("numId"))] = a.get(W("val"))

    def _list_marker(self, num_id: str, ilvl: str) -> str:
        if num_id in (None, "0"):
            return ""
        fmt = self.abs_fmt.get((self.num_abs.get(num_id), ilvl), "bullet")
        lvl = int(ilvl or 0)
        key = (self.num_abs.get(num_id, num_id), lvl)
        self.counters[key] = self.counters.get(key, 0) + 1
        for k in list(self.counters):  # a higher-level item restarts deeper levels
            if k[0] == key[0] and k[1] > lvl:
                del self.counters[k]
        n = self.counters[key]
        indent = "  " * lvl
        if fmt == "bullet":
            return indent + "- "
        if fmt in ("lowerLetter", "upperLetter"):
            c = chr(ord("a") + (n - 1) % 26)
            return indent + (c.upper() if fmt == "upperLetter" else c) + ". "
        if fmt in ("lowerRoman", "upperRoman"):
            r = _roman(n)
            return indent + (r if fmt == "upperRoman" else r.lower()) + ". "
        return indent + f"{n}. "

    # --- body ----------------------------------------------------------------
    def _render_body(self):
        body = self.doc.find(W("body"))
        for el in self._blocks(body):
            if el.tag == W("p"):
                self._paragraph(el)
            elif el.tag == W("tbl"):
                self._table(el)

    def _blocks(self, container):
        """Block-level children, unwrapping content controls / custom XML."""
        for el in container:
            if el.tag in (W("p"), W("tbl")):
                yield el
            elif el.tag in (W("sdt"),):
                c = el.find(W("sdtContent"))
                if c is not None:
                    yield from self._blocks(c)
            elif el.tag == W("customXml"):
                yield from self._blocks(el)

    def _paragraph(self, p, in_table=False) -> str:
        self.para += 1
        n = self.para
        ppr = p.find(W("pPr"))
        sid = None
        numpr = None
        if ppr is not None:
            ps = ppr.find(W("pStyle"))
            sid = ps.get(W("val")) if ps is not None else None
            numpr = ppr.find(W("numPr"))
        style = display_style(self.style_name.get(sid, sid or "Normal")) if sid else "Normal"
        numpr = numpr if numpr is not None else self._style_numpr(sid)
        pieces: list[Piece] = []
        self._inline(p, pieces, n)
        text = compose(pieces).strip()
        level = self.heading_level_of(sid)
        has_image = p.find(f".//{W('drawing')}") is not None or p.find(f".//{W('pict')}") is not None
        abs_id = None
        if numpr is not None:
            nid = numpr.find(W("numId"))
            abs_id = self.num_abs.get(nid.get(W("val"))) if nid is not None else None
        plain = "".join(pc.text for pc in pieces if not pc.raw and pc.change != "del")
        hidden = ppr is not None and on(ppr.find(W("rPr")), "vanish")
        self.para_meta.append({"n": n, "sid": sid or "Normal", "text": bool(text), "in_table": in_table,
                               "list": self._num_kind(numpr), "abs": abs_id, "heading": level, "image": has_image,
                               "plain": plain, "style": style, "hidden": hidden,
                               "section_break": ppr is not None and ppr.find(W("sectPr")) is not None})
        if in_table:
            return text
        if not text:
            return ""
        prefix = ""
        low = style.lower()
        m = re.match(r"heading (\d)", low)
        if m or level:
            lvl = int(m.group(1)) if m else level
            prefix = "#" * lvl + " "
            self.heading_level[n] = lvl
        elif low == "title":
            self.heading_level[n] = 0
            prefix = "# "
        elif numpr is not None:
            nid = numpr.find(W("numId"))
            lv = numpr.find(W("ilvl"))
            prefix = self._list_marker(nid.get(W("val")) if nid is not None else None,
                                       lv.get(W("val")) if lv is not None else "0")
        label = f"[p{n}]" if style == "Normal" or m or numpr is not None else f"[p{n} {style}]"
        self.lines.append((n, f"{label} {prefix}{text}"))
        return text

    def _unwrapped(self, parent, name):
        """Rows or cells, including those a content control (w:sdt) wraps - cover pages bind
        document properties to a cell this way."""
        for el in parent:
            if el.tag == W(name):
                yield el
            elif el.tag == W("sdt"):
                c = el.find(W("sdtContent"))
                if c is not None:
                    yield from self._unwrapped(c, name)

    def _table(self, tbl, nested=False) -> list[list[str]]:
        if not nested:
            self.tables += 1
            ts = tbl.find(f"{W('tblPr')}/{W('tblStyle')}")
            self.table_styles.append(ts.get(W("val")) if ts is not None else "TableNormal")
        first = self.para + 1
        rows = []
        for tr in self._unwrapped(tbl, "tr"):
            row = []
            for tc in self._unwrapped(tr, "tc"):
                tcpr = tc.find(W("tcPr"))
                span = 1
                vmerge = False
                if tcpr is not None:
                    gs = tcpr.find(W("gridSpan"))
                    span = int(gs.get(W("val"), "1")) if gs is not None else 1
                    vm = tcpr.find(W("vMerge"))
                    vmerge = vm is not None and vm.get(W("val"), "continue") == "continue"
                parts = []
                for el in self._blocks(tc):
                    if el.tag == W("p"):
                        t = self._paragraph(el, in_table=True)
                        if t:
                            parts.append(t)
                    else:
                        sub = self._table(el, nested=True)
                        parts.append("; ".join(" | ".join(r) for r in sub))
                cell = " / ".join(parts).replace("|", "\\|")
                row.append("^" if vmerge and not cell else cell)
                row.extend([""] * (span - 1))
            self.para += 1  # end-of-row mark is a paragraph in Word's numbering
            self.para_meta.append({"n": self.para, "sid": "", "text": False, "in_table": True, "list": None,
                                   "abs": None, "heading": None, "image": False, "plain": "", "style": "",
                                   "hidden": False, "section_break": False, "row_end": True})
            rows.append(row)
        if not nested and rows:
            cols = max(len(r) for r in rows)
            rows = [r + [""] * (cols - len(r)) for r in rows]
            out = [f"[p{first}-p{self.para}] table {self.tables} ({len(rows)} rows x {cols} columns):",
                   "| " + " | ".join(rows[0]) + " |", "|" + " --- |" * cols]
            out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
            self.lines.append((first, "\n".join(out)))
        return rows

    # --- inline content --------------------------------------------------------
    def _inline(self, el, out: list, para: int, change="", link=""):
        for c in el:
            t = c.tag
            if t == W("r"):
                self._run(c, out, para, change, link)
            elif t == W("hyperlink"):
                rid = c.get(q("r:id"))
                url = self.rels.get(rid, ("", ""))[1] if rid else "#" + c.get(W("anchor"), "")
                self._inline(c, out, para, change, url)
            elif t in (W("ins"), W("moveTo")):
                self.changes["ins"] += 1
                self._inline(c, out, para, "ins", link)
            elif t in (W("del"), W("moveFrom")):
                self.changes["del"] += 1
                self._inline(c, out, para, "del", link)
            elif t in (W("smartTag"), W("customXml"), W("fldSimple")):
                self._inline(c, out, para, change, link)
            elif t == W("sdt"):
                sc = c.find(W("sdtContent"))
                if sc is not None:
                    self._inline(sc, out, para, change, link)
            elif t == W("commentRangeStart"):
                cid = c.get(W("id"))
                self.open_comments.add(cid)
                self.comment_anchor.setdefault(cid, [])
                self.comment_para.setdefault(cid, para)
            elif t == W("commentRangeEnd"):
                self.open_comments.discard(c.get(W("id")))

    def _run(self, r, out: list, para: int, change: str, link: str):
        rpr = r.find(W("rPr"))
        rs = rpr.find(W("rStyle")) if rpr is not None else None
        if rs is not None:
            cs = rs.get(W("val"))
            self.char_styles[cs] = self.char_styles.get(cs, 0) + 1
        fmt = dict(bold=on(rpr, "b"), italic=on(rpr, "i"), strike=on(rpr, "strike") or on(rpr, "dstrike"),
                   change=change, link=link)
        for c in r:
            t = c.tag
            if t == W("fldChar"):
                kind = c.get(W("fldCharType"))
                if kind == "begin":
                    self.field_stack.append("code")
                elif kind == "separate" and self.field_stack:
                    self.field_stack[-1] = "result"
                elif kind == "end" and self.field_stack:
                    self.field_stack.pop()
                continue
            if self.field_stack and self.field_stack[-1] == "code":
                continue
            if t in (W("t"), W("delText")):
                self._emit(out, Piece(c.text or "", **fmt))
            elif t == W("tab"):
                self._emit(out, Piece("\t", **fmt))
            elif t in (W("br"), W("cr")):
                self._emit(out, Piece(" [page break] " if c.get(W("type")) == "page" else "\n", raw=True))
            elif t == W("noBreakHyphen"):
                self._emit(out, Piece("-", **fmt))
            elif t == W("drawing"):
                self._emit(out, Piece(self._drawing(c, para), raw=True))
            elif t == W("pict"):
                self._emit(out, Piece(self._vml(c, para), raw=True))
            elif t == W("footnoteReference"):
                self.footnote_refs.append(c.get(W("id")))
                self._emit(out, Piece(f"[^{c.get(W('id'))}]", raw=True))
            elif t == W("endnoteReference"):
                self._emit(out, Piece(f"[^e{c.get(W('id'))}]", raw=True))
            elif t == W("commentReference"):
                cid = c.get(W("id"))
                if cid not in self.comment_order:
                    self.comment_order.append(cid)
                    self.comment_para.setdefault(cid, para)
                    if self._comment_parent(cid) is None:  # replies share the parent's marker
                        self._emit(out, Piece(f"[c{len(self.comment_order)}]", raw=True))

    def _emit(self, out: list, piece: Piece):
        if not piece.raw:
            for cid in self.open_comments:
                self.comment_anchor[cid].append(piece.text)
        out.append(piece)

    def _drawing(self, d, para: int) -> str:
        holder = d[0] if len(d) else d
        floating = holder.tag == q("wp:anchor")
        ext = holder.find(q("wp:extent"))
        w_cm = int(ext.get("cx", 0)) / EMU_PER_CM if ext is not None else 0
        h_cm = int(ext.get("cy", 0)) / EMU_PER_CM if ext is not None else 0
        pr = holder.find(q("wp:docPr"))
        alt = (pr.get("descr") or pr.get("title") or "") if pr is not None else ""
        name = pr.get("name", "") if pr is not None else ""
        blip = d.find(f".//{q('a:blip')}")
        gd = d.find(f".//{q('a:graphicData')}")
        uri = gd.get("uri", "") if gd is not None else ""
        if blip is not None and blip.get(q("r:embed")):
            target = self.rels.get(blip.get(q("r:embed")), ("", ""))[1]
            return self._image(target, alt, name, w_cm, h_cm, floating, para)
        txt = [compose(self._collect(p)) for p in d.iter(W("p"))]
        if txt:
            return "[text box: " + " / ".join(t for t in txt if t) + "]"
        if uri.endswith("/chart"):
            return f"[chart: {name or alt}]"
        if "diagram" in uri:
            return f"[SmartArt: {name or alt}]"
        return f"[shape: {name or alt}]"

    def _vml(self, pict, para: int) -> str:
        img = pict.find(f".//{q('v:imagedata')}")
        if img is not None and img.get(q("r:id")):
            target = self.rels.get(img.get(q("r:id")), ("", ""))[1]
            return self._image(target, "", "", 0, 0, False, para)
        txt = [compose(self._collect(p)) for p in pict.iter(W("p"))]
        return "[text box: " + " / ".join(t for t in txt if t) + "]" if txt else "[shape]"

    def _collect(self, p) -> list:
        out: list = []
        saved = (self.field_stack, self.open_comments)
        self.field_stack, self.open_comments = [], set()
        try:
            self._inline(p, out, self.para)
        finally:
            self.field_stack, self.open_comments = saved
        return out

    def _image(self, part, alt, name, w_cm, h_cm, floating, para) -> str:
        ctype = self.pkg.binary.get(part, ("", b""))[0]
        k = len(self.images) + 1
        if not floating:
            self.inline_images += 1
        self.images.append({"n": k, "part": part, "content_type": ctype, "alt_text": alt, "name": name,
                            "width_cm": round(w_cm, 2), "height_cm": round(h_cm, 2),
                            "floating": floating, "paragraph": para,
                            "inline_index": None if floating else self.inline_images})
        desc = alt or name or part.rsplit("/", 1)[-1]
        size = f", {w_cm:.1f}x{h_cm:.1f} cm" if w_cm else ""
        return f"![image {k}: {desc}{size}{', floating' if floating else ''}]"

    # --- comments / notes / properties -----------------------------------------
    def _comments_ext(self):
        if hasattr(self, "_cx"):
            return self._cx
        para_of = {}   # comment id -> its last paragraph's w14:paraId
        cm = self.pkg.parts.get("/word/comments.xml")
        if cm is not None:
            for c in cm.findall(W("comment")):
                ps = c.findall(W("p"))
                if ps:
                    para_of[c.get(W("id"))] = ps[-1].get(q("w14:paraId"))
        ext = {}
        cx = self.pkg.parts.get("/word/commentsExtended.xml")
        if cx is not None:
            for e in cx.findall(q("w15:commentEx")):
                ext[e.get(q("w15:paraId"))] = (e.get(q("w15:paraIdParent")), e.get(q("w15:done")) == "1")
        by_para = {v: k for k, v in para_of.items()}
        out = {}
        for cid, pid in para_of.items():
            parent_pid, done = ext.get(pid, (None, False))
            out[cid] = (by_para.get(parent_pid) if parent_pid else None, done)
        self._cx = out
        return out

    def _comment_parent(self, cid):
        return self._comments_ext().get(cid, (None, False))[0]

    def comments(self) -> list[dict]:
        """Comments in Word's Comments(n) order (n = reference order), threaded via parent."""
        cm = self.pkg.parts.get("/word/comments.xml")
        if cm is None:
            return []
        info = {}
        for c in cm.findall(W("comment")):
            cid = c.get(W("id"))
            text = "\n".join(compose(self._collect(p)) for p in c.findall(W("p"))).strip()
            info[cid] = {"author": c.get(W("author"), ""), "date": c.get(W("date"), ""), "text": text}
        order = self.comment_order + [cid for cid in info if cid not in self.comment_order]
        n_of = {cid: i + 1 for i, cid in enumerate(order)}
        ext = self._comments_ext()
        out = []
        for cid in order:
            if cid not in info:
                continue
            parent, done = ext.get(cid, (None, False))
            out.append({"n": n_of[cid], **info[cid], "resolved": done,
                        "reply_to": n_of.get(parent) if parent else None,
                        "on": "".join(self.comment_anchor.get(cid, [])).strip(),
                        "paragraph": self.comment_para.get(cid)})
        return out

    def notes(self, part: str, tag: str) -> list[tuple[str, str]]:
        el = self.pkg.parts.get(part)
        if el is None:
            return []
        out = []
        for f in el.findall(W(tag)):
            if f.get(W("type")) in ("separator", "continuationSeparator", "continuationNotice"):
                continue
            txt = " ".join(compose(self._collect(p)) for p in f.findall(W("p"))).strip()
            out.append((f.get(W("id")), txt))
        return out

    def headers_footers(self) -> dict:
        out = {"header": [], "footer": []}
        for _, (kind, target) in self.rels.items():
            if kind in out and target in self.pkg.parts:
                txt = " / ".join(t for t in (compose(self._collect(p)).strip()
                                             for p in self.pkg.parts[target].iter(W("p"))) if t)
                if txt and txt not in out[kind]:
                    out[kind].append(txt)
        return out

    def properties(self) -> dict:
        core = self.pkg.parts.get("/docProps/core.xml")
        out = {}
        if core is not None:
            for key, tag in (("title", "dc:title"), ("subject", "dc:subject"), ("author", "dc:creator"),
                             ("last_modified_by", "cp:lastModifiedBy"), ("created", "dcterms:created"),
                             ("modified", "dcterms:modified")):
                el = core.find(q(tag))
                if el is not None and el.text:
                    out[key] = el.text
        return out

    def image_bytes(self, n: int) -> tuple[dict, str, bytes]:
        if not 1 <= n <= len(self.images):
            raise IndexError(f"image {n} does not exist (document has {len(self.images)})")
        meta = self.images[n - 1]
        ctype, data = self.pkg.binary.get(meta["part"], ("", ""))
        if not data:
            raise IndexError(f"image {n} has no embedded data (linked picture?)")
        return meta, ctype, base64.b64decode(data)

    # --- house style -----------------------------------------------------------------
    def house_style(self) -> dict:
        """Which style the document itself uses for each kind of content, so new content matches it.
        Usage in the document wins; then styles the template defines but nobody used yet; then Word's
        built-ins. Values: {"id": styleId, "name": style name, "source": "used" | "template" | "builtin"}."""
        name = self.style_name

        def pick(counter: dict, source="used"):
            if not counter:
                return None
            sid = max(counter, key=counter.get)
            return {"id": sid, "name": name.get(sid, sid), "source": source, "custom": sid in self.style_custom}

        def count(pred):
            c: dict[str, int] = {}
            for m in self.para_meta:
                if m["text"] and pred(m):
                    c[m["sid"]] = c.get(m["sid"], 0) + 1
            return c

        def defined(pred, typ="paragraph"):
            """Template styles (in styles.xml, custom) matching pred, by name."""
            c = {sid: 1 for sid, t in self.style_type.items() if t == typ and sid in self.style_custom and pred(sid)}
            return pick(c, "template") if len(c) else None

        def builtin(n, sid):
            return {"id": sid, "name": n, "source": "builtin", "custom": False}

        role_name = lambda pat: (lambda sid: re.search(pat, name.get(sid, ""), re.I) is not None)  # noqa: E731
        out: dict = {}
        for lvl in range(1, 7):
            out[f"heading{lvl}"] = (pick(count(lambda m, l=lvl: m["heading"] == l and not m["in_table"]))
                                    or defined(lambda sid, l=lvl: self.heading_level_of(sid) == l)
                                    or builtin(f"heading {lvl}", f"Heading{lvl}"))
        out["title"] = (pick(count(lambda m: name.get(m["sid"], "").lower() == "title"))
                        or builtin("Title", "Title"))
        plain = lambda m: (not m["in_table"] and not m["heading"] and not m["list"] and not m["image"]  # noqa: E731
                           and not re.search(r"caption|toc|title|quote|code|source|listing", name.get(m["sid"], ""), re.I))
        used_body = count(plain)
        out["body"] = (pick({k: v for k, v in used_body.items() if k != "Normal"})
                       or defined(role_name(r"body"))
                       or pick(used_body) or builtin("Normal", "Normal"))
        for kind in ("bullet", "number"):
            out[kind] = (pick(count(lambda m, k=kind: m["list"] == k and not m["in_table"]))
                         or builtin("List Paragraph", "ListParagraph"))
            abs_use: dict[str, int] = {}
            for m in self.para_meta:
                if m["list"] == kind and m["abs"] is not None:
                    abs_use[m["abs"]] = abs_use.get(m["abs"], 0) + 1
            if abs_use:  # the document's own bullets / numbering format, copied into new lists
                out[kind]["abstract_xml"] = self.abstract_xml(max(abs_use, key=abs_use.get))
        out["table"] = (pick({s: self.table_styles.count(s) for s in self.table_styles if s != "TableNormal"})
                        or defined(lambda sid: True, "table") or builtin("Table Grid", "TableGrid"))
        out["table_text"] = pick(count(lambda m: m["in_table"] and m["sid"] != "Normal"))
        out["code"] = (pick(count(lambda m: role_name(r"code|source|listing|preformat|snippet")(m["sid"])))
                       or defined(role_name(r"code|source|listing|preformat|snippet")))
        code_chars = {s: c for s, c in self.char_styles.items() if role_name(r"code|mono|consol|source")(s)}
        out["inline_code"] = pick(code_chars) or defined(role_name(r"code|mono|consol|source"), "character")
        out["quote"] = (pick(count(lambda m: role_name(r"quote")(m["sid"])))
                        or builtin("Quote", "Quote"))
        out["caption"] = (pick(count(lambda m: role_name(r"caption")(m["sid"]))) or builtin("caption", "Caption"))
        out["figure"] = pick({k: v for k, v in count(lambda m: m["image"] and not m["in_table"]).items() if k != "Normal"})
        out["caption_label"] = max(self.seq_labels, key=self.seq_labels.get) if self.seq_labels else "Figure"
        return {k: v for k, v in out.items() if v is not None}

    def abstract_xml(self, abs_id: str) -> str | None:
        nm = self.pkg.parts.get("/word/numbering.xml")
        if nm is None:
            return None
        for a in nm.findall(W("abstractNum")):
            if a.get(W("abstractNumId")) == abs_id:
                return ET.tostring(a, encoding="unicode")
        return None

    def style_catalog(self) -> list[dict]:
        """Styles defined in the document (not latent ones), with use counts and their look."""
        uses: dict[str, int] = {}
        for m in self.para_meta:
            if m["text"]:
                uses[m["sid"]] = uses.get(m["sid"], 0) + 1
        for s in self.table_styles:
            uses[s] = uses.get(s, 0) + 1
        for s, c in self.char_styles.items():
            uses[s] = uses.get(s, 0) + c
        out = []
        for sid, nm in self.style_name.items():
            t = self.style_type.get(sid, "paragraph")
            if t == "numbering" or (not uses.get(sid) and sid not in self.style_custom):
                continue
            item = {"name": nm, "type": t, "used": uses.get(sid, 0), "custom": sid in self.style_custom}
            lvl = self.heading_level_of(sid)
            if lvl:
                item["heading_level"] = lvl
            if sid in self.style_look:
                item["look"] = self.style_look[sid]
            out.append(item)
        out.sort(key=lambda x: (-x["used"], x["name"]))
        return out

    # --- output ------------------------------------------------------------------
    def render(self, start_paragraph: int = 1, max_chars: int = 60000) -> tuple[str, int | None]:
        """Body text from start_paragraph; returns (text, next_paragraph or None when complete)."""
        out, size = [], 0
        for n, line in self.lines:
            if n < start_paragraph:
                continue
            if size + len(line) > max_chars and out:
                return "\n".join(out), n
            out.append(line)
            size += len(line) + 1
        return "\n".join(out), None


def display_style(name: str) -> str:
    """Word's display name for a style: styles.xml spells built-ins in lowercase ("heading 1",
    "caption", "toc 2"); Word shows "Heading 1", "Caption", "TOC 2"."""
    if not name or not name[0].islower():
        return name
    m = re.match(r"toc (\d)$", name)
    if m:
        return f"TOC {m.group(1)}"
    return name[0].upper() + name[1:]


def compose(pieces: list) -> str:
    out = []
    for p in pieces:
        if p.raw:
            out.append(p.text)
            continue
        t = p.text
        if not t:
            continue
        core = t.strip()
        if core:
            lead, trail = t[: len(t) - len(t.lstrip())], t[len(t.rstrip()):]
            if p.bold and p.italic:
                core = f"***{core}***"
            elif p.bold:
                core = f"**{core}**"
            elif p.italic:
                core = f"*{core}*"
            if p.strike:
                core = f"~~{core}~~"
            if p.link and not p.link.startswith("#"):
                core = f"[{core}]({p.link})"
            t = lead + core + trail
        if p.change == "ins":
            t = "{+" + t + "+}"
        elif p.change == "del":
            t = "{-" + t + "-}"
        out.append(t)
    s = "".join(out)
    # merge adjacent identical markers produced by run splits: **a****b** -> **ab**
    s = re.sub(r"\*\*\*\*", "", s)
    s = s.replace("+}{+", "").replace("-}{-", "")
    return s


def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out
