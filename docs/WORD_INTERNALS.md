# Word internals for word-mcp maintainers

Companion documents: [REA_ANALYSIS.md](REA_ANALYSIS.md) (how Word was reverse engineered, with evidence,
and how to reproduce it without REA) and [WORD_API_REFERENCE.md](WORD_API_REFERENCE.md) (generated
signatures of the Word API the server uses).

Everything this project learned by reverse engineering and measuring Microsoft Word, written so that an
agent which hits an error while using or changing word-mcp can find the cause and fix it without
rediscovering it. Facts are from Word 16 (Microsoft 365, builds 16.0.14332 and 16.0.20326) on Windows
11 with pywin32; statements marked *(unverified)* are general Word knowledge not re-tested here.

Contents:

1. [How Word is built (binaries)](#1-how-word-is-built)
2. [The automation surface (COM, type library)](#2-the-automation-surface)
3. [Talking to Word from Python (threading, errors, speed)](#3-talking-to-word-from-python)
4. [The text model: characters, paragraphs, numbering](#4-the-text-model)
5. [The object model pieces word-mcp uses](#5-object-model-pieces)
6. [The file format: WordOpenXML (flat OPC)](#6-wordopenxml)
7. [Writing with Range.InsertXML](#7-writing-with-rangeinsertxml)
8. [Styles, house style and lists](#8-styles-and-lists)
9. [Comments, revisions, fields, pictures, pages](#9-comments-revisions-fields-pictures-pages)
10. [Word behaviours that surprise automation](#10-surprising-behaviours)
11. [word-mcp architecture and invariants](#11-word-mcp-architecture)
12. [Troubleshooting playbook: symptom -> cause -> fix](#12-troubleshooting-playbook)
13. [How to fix and verify a bug](#13-how-to-fix-and-verify-a-bug)
14. [Re-deriving all of this (REA and probes)](#14-re-deriving-this)

---

## 1. How Word is built

Install folder (Click-to-Run): `C:\Program Files\Microsoft Office\root\Office16\`.

| File | Size | Role (how we know) |
|---|---|---|
| `WINWORD.EXE` | 1.6 MB | **Launcher only.** REA `inspect_pe_resources`: its resources are 225 icons, 13 icon groups, a manifest, version info - no UI or document code. Its import table has no static link to Word code (only ADVAPI32 registry calls, KERNEL32, the C runtime and `AppVIsvSubsystems64.dll`, the Click-to-Run virtualization shim). It contains the UTF-16 string `wwlib` and the ASCII strings `LoadLibrary`, `GetProcAddress`, `FMain`: it loads `WWLIB.DLL` at run time and calls its `FMain` export. |
| `WWLIB.DLL` | 45.8 MB | **All of Word**: document model, layout, rendering (imports GDI32, gdiplus), file formats, OLE/COM (ole32, OLEAUT32), Office art (`oart.dll`). Plain-C exports: `FMain` (Word's main), `DllMain`, `DllGetClassObject` (COM class factory: this is how `Word.Application` objects are created), `DllCanUnloadNow`, `DllGetLCID`; plus 14 C++ exports for `WordMailReact` (Outlook/mail editing). Sections: `.text .rdata .data .pdata .didat _RDATA .rsrc .reloc` (`.didat` = delay-loaded imports, so many Office DLLs load lazily). Automation member names (`InsertXML`, `WordOpenXML`...) do **not** occur as strings in the DLL - names are resolved through the type library. |
| `MSWORD.OLB` | 0.95 MB | **The automation contract.** A PE image whose only real resource is `TYPELIB` (931,600 bytes). Type library GUID `{00020905-0000-0000-C000-000000000046}`, version **8.7**, registered under `HKCR\TypeLib\{00020905-...}\8.7\0\win64`. |
| `Wordcnvpxy.cnv` | 50 KB | Converter proxy for legacy text converters *(role inferred from name; unverified)*. |

Why this matters: the 45 MB engine is far too large for interactive decompilation (REA/Ghidra time out
on it), and nothing an MCP needs is hidden there. Everything Word lets programs do is described by the
type library and reachable through COM; everything about a document's structure is visible in its
OOXML. Those two are the sources of truth this document builds on.

## 2. The automation surface

`dump_typelib.ps1` decodes `MSWORD.OLB` with `LoadTypeLibEx` into `word-api.json` (built on first use of
`word_api_search`; about 2 s). Contents for 16.0.14332:

- 767 types: **374 dispatch interfaces**, **385 enums** (3,754 constants), 7 coclasses, 1 interface;
- 10,445 visible members, of which 4,072 methods.

Reading it:

- Interfaces are prefixed: `_Application`, `_Document`, `_Font`, `_ParagraphFormat`; `Range`,
  `Selection`, `Paragraphs`, `Comments`, etc. are dispatch interfaces of the same name.
- Parameters marked optional are `VARIANT*`; positional order matters when calling late-bound (pywin32
  dynamic dispatch passes positionally).
- Enum values are what properties return: e.g. `WdBuiltinStyle` (`wdStyleHeading1 = -2`,
  `wdStyleNormal = -1`, `wdStyleTitle = -63`, `wdStyleCaption = -35`, `wdStyleListParagraph = -180`,
  `wdStyleQuote = -181`, `wdStyleHyperlink = -86`), `WdBreakType` (`wdPageBreak = 7`,
  `wdSectionBreakNextPage = 2`), `WdListNumberStyle` (`wdListNumberStyleBullet = 23`).
- Look anything up with the `word_api_search` tool: `"Range"` lists Range's members,
  `"InsertCaption"` finds a member, `"WdBreakType"` lists an enum.

## 3. Talking to Word from Python

**Attach, don't launch.** `win32com.client.GetActiveObject("Word.Application")` attaches to the running
Word (raises `MK_E_UNAVAILABLE 0x800401E3` if none; then `Dispatch("Word.Application")` starts one and
`Visible = True` shows it). word-mcp wraps the result in `win32com.client.dynamic.Dispatch` (late
binding).

**Late vs early binding.** Measured: equal speed (both about 1.5 ms per property hop). Early binding
(`gencache.EnsureModule(GUID, 0, 8, 7)`, 1 s to generate) gives constants and call-by-dispid but writes
a cache; late binding is simpler and is what the server uses. Late-bound pitfalls seen:
`Application.CentimetersToPoints(2.0)` fails with `0x80004005` - do unit math in Python
(1 cm = 72/2.54 pt).

**One STA thread.** All COM objects are created and used on one dedicated thread
(`ComWorker` in `word_mcp_server.py`, `pythoncom.CoInitialize()` once). Never pass a COM object to
another thread. The MCP event loop only queues closures to that thread.

**Cost model.** Word is out-of-process: every property hop is an RPC.

| Call | Measured |
|---|---|
| Held object property (`doc.Name`) | 0.4-0.6 ms |
| Chained new objects (`doc.Range(0,1).Text`, `Paragraphs.Count`) | 2-5 ms |
| `ScreenUpdating = False/True` | 8 ms+ (the repaint) |
| `UndoRecord.StartCustomRecord/EndCustomRecord` | 0.3 ms |
| `Content.Text` of 500 paragraphs | 6 ms |
| `Document.WordOpenXML` (small doc / 40 screenshots) | 23 ms / ~150 ms |
| `Range.InsertXML` of a full section / of 40 sections + 40 images | 0.07-0.14 s / 0.54 s |
| Find/Replace with formatting over 500 matches | 0.4 s (vs 13 s formatting match by match) |
| `Pages(n).EnhMetaFileBits` (render a page) | 0.4-0.7 s |
| Walking `Paragraphs` one by one | ~8 ms per paragraph |

Rule: never loop over paragraphs/cells/matches through COM; read in bulk (`Content.Text`,
`WordOpenXML`) and write in bulk (`InsertXML`, `ConvertToTable`, Find/Replace).

**Errors you will see** (HRESULT and meaning):

| HRESULT | Meaning | Handling |
|---|---|---|
| `0x80010001 RPC_E_CALL_REJECTED`, `0x8001010A RPC_E_SERVERCALL_RETRYLATER` | Word's UI thread is busy: a dialog, a menu, a context menu or a scrollbar drag is open, or it is mid-operation | retry up to 5 s, then "Word is busy" |
| `0x80010108 RPC_E_DISCONNECTED`, `0x800706BA`, `0x800401FD` | Word was closed/restarted | drop the cached Application, reattach once |
| `0x800401E3 MK_E_UNAVAILABLE` | Word not running | start it |
| `0x80020009 DISP_E_EXCEPTION` | Word raised an error; the text is in excepinfo[2] and the Word error number in scode `0x800A0000 + n` | `com_message()` surfaces excepinfo text |
| scode `0x800A1735` (5941) | "The requested member of the collection does not exist" (bad index/name) | |
| "This command is not available" (Word error 4605, scode `0x800A11FD` *(code unverified; message observed)*) | e.g. `DeleteAllComments` with no comments; `TrackRevisions` while a comment card is selected | read with `safe()`, use alternatives |
| `0x800A1200` | OLE error setting some `ListLevel` properties late-bound (`NumberStyle`) | use `ApplyBulletDefault` + `NumberFormat` |
| `AttributeError: <unknown>.Name` (pywin32) | Word handed back an object it could not describe yet (document mid-save/sync) | transient: retry, then explain |
| `0x80004005 E_FAIL` | generic; seen on late-bound unit conversion | avoid the call |

Set `WORD_MCP_LOG=<file>` to get every failed call with its Python trace.

## 4. The text model

`Range.Text` / `Document.Content.Text` use these control characters:

| Char | Meaning |
|---|---|
| `\r` (13) | paragraph mark |
| `\r\x07` | end of a table cell **and** end of a table row (both are paragraphs in Word's numbering) |
| `\x0b` (11) | manual line break (Shift+Enter) |
| `\x0c` (12) | page break **or** section break - indistinguishable in text |
| `\x13 \x14 \x15` | field begin / separator / end (codes shown only when field codes are visible) |
| `\x1e` / `\x1f` | non-breaking / optional hyphen |
| `/` | an inline picture shows up as `/` in `Range.Text` on these builds (observed) |
| hidden text | omitted from `Range.Text` by default (a hidden paragraph reads as empty) |

### Paragraph numbering (the rule that matters most)

Word's `Document.Paragraphs` collection counts:

- every body paragraph, including empty ones and paragraphs holding only a picture;
- every **table cell** paragraph, plus **one paragraph per row end**;
- a paragraph that ends in a **section break** (its mark is the break) as one paragraph;
- text-box contents, headers, footers, comments and footnotes are **other stories**, not counted.

**`Content.Text.split("\r")` is NOT Word's numbering**: a section break ends a paragraph without a
`\r`, so the split merges the section-break paragraph with the next one and falls one behind per
section break. Measured on the fixture: Word 76 paragraphs, split 74, WordOpenXML parse 76. This
mismatch caused real damage (anchored inserts in the wrong place, deletes removing a picture and a
comment). word-mcp therefore numbers paragraphs from the WordOpenXML parse (`paragraph_map()`), which
counts exactly like Word, falls back to walking `Paragraphs` if the counts ever differ, and verifies
every target with `d.Paragraphs(n).Range.Text` before editing. `split_texts()` survives only as a cheap
change-detection baseline and is translated to Word numbers when reported.

Position <-> paragraph: `doc.Range(0, pos + 1).Paragraphs.Count` is the paragraph containing `pos`
(the `+1` matters when `pos` is the first character of a paragraph).

## 5. Object model pieces

- `Application`: `Documents`, `ActiveDocument`, `Selection`, `DisplayAlerts` (0 = none; set it around
  `Open`/`SaveAs2` so no dialog blocks automation), `ScreenUpdating`, `UndoRecord`, `Options`,
  `CaptionLabels`.
- `Documents.Add([Template], ..., [Visible])` - `Visible=False` gives a hidden document that still has
  a window and pages (good for probes). `Documents.Open(path_or_url, ConfirmConversions, ReadOnly,
  AddToRecentFiles, ...)` accepts http(s) URLs for SharePoint/OneDrive.
- `Document`: `Content`, `Paragraphs`, `Range(start, end)`, `Tables`, `InlineShapes` (pictures in the
  text flow), `Shapes` (floating), `Comments`, `Revisions`, `TrackRevisions`, `Fields`,
  `TablesOfContents`, `Sections`, `Styles`, `ListTemplates`, `PageSetup`, `Windows`, `WordOpenXML`,
  `Saved`, `Path` (empty = never saved; an `https://` path = cloud), `FullName`, `AutoSaveOn`
  (SharePoint/OneDrive; local files raise), `CopyStylesFromTemplate(path)`, `Undo(n)`/`Redo(n)`,
  `ComputeStatistics(2 = pages, 0 = words)` (forces pagination; skip on huge documents).
- `Range`: `Start/End`, `Text`, `InsertBefore/InsertAfter`, `InsertParagraphAfter`, `Delete`,
  `Style`, `Font`, `ParagraphFormat`, `ListFormat` (`ApplyListTemplate`, `RemoveNumbers`,
  `ListString`), `Find`, `InsertXML`, `WordOpenXML`, `ConvertToTable`, `InsertCaption(label, title,
  autotext, position)`, `Information(12 = wdWithInTable, 2 = section number)`, `Fields.Update`,
  `InlineShapes`, `Comments`, `Tables`.
- `Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
  MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace, ...)` - positional order as listed;
  `Wrap` 0 = stop, 1 = continue; `Replace` 1 = one, 2 = all. `"^&"` as ReplaceWith re-inserts the found
  text, which lets `Replacement.Font` / `Replacement.Style` format every match in one call. Highlight in
  a replacement uses `Options.DefaultHighlightColorIndex` (save and restore it).
- `Windows(1).Panes(1).Pages(n).EnhMetaFileBits` - a page as EMF bytes; needs Print Layout view
  (`View.Type = 3`). Pillow opens EMF on Windows: `Image.open(...).load(dpi=96)`.
- `UndoRecord.StartCustomRecord("name")` / `EndCustomRecord()` - groups everything in between into one
  Ctrl+Z step (the undo history is shared with the user's own typing).
- Colours are **BGR** integers (`#RRGGBB` -> `r | g<<8 | b<<16`); booleans read back as `-1`/`0`;
  mixed formatting reads back as `9999999` (`wdUndefined`).

## 6. WordOpenXML

`Document.WordOpenXML` (and `Range.WordOpenXML`) returns the whole document as **flat OPC**: one XML
string `<pkg:package>` with `<pkg:part pkg:name="/word/document.xml">` parts holding `<pkg:xmlData>` or
base64 `<pkg:binaryData>` (images). Parts seen: `/_rels/.rels`, `/word/document.xml`,
`/word/_rels/document.xml.rels`, `/word/styles.xml`, `/word/numbering.xml`, `/word/comments.xml`,
`/word/commentsExtended.xml`, `/word/commentsIds.xml`, `/word/commentsExtensible.xml`,
`/word/people.xml`, `/word/media/imageN.png`, `/word/theme/theme1.xml`, `/word/settings.xml`,
`/word/fontTable.xml`, `/word/webSettings.xml`, header/footer/footnote parts, `/docProps/core.xml`,
`/docProps/app.xml`. It reflects **unsaved** in-memory state (the user's typing included).

Structures `docx_context.py` relies on:

- Body: `w:body` children `w:p` (paragraph), `w:tbl` (table: `w:tr` rows, `w:tc` cells with their own
  `w:p`/`w:tbl`), `w:sdt` (content control: unwrap `w:sdtContent`), `w:customXml`. A paragraph's
  `w:pPr/w:sectPr` marks the **section break** that ends it. The last `w:sectPr` (body child) is the
  final section.
- Runs: `w:r` with `w:rPr` (`w:b`, `w:i`, `w:strike`, `w:rStyle`, `w:vanish` = hidden) and content
  `w:t`, `w:tab`, `w:br` (`w:type="page"`), `w:drawing`, `w:pict` (VML), `w:footnoteReference`,
  `w:commentReference`, `w:fldChar` (begin/separate/end) + `w:instrText` (field code), `w:fldSimple`.
- Tracked changes: `w:ins`, `w:del` (deleted text in `w:delText`), `w:moveTo`, `w:moveFrom`.
- Links: `w:hyperlink r:id` -> relationship target, or `w:anchor` (bookmark).
- Pictures: `w:drawing/wp:inline` (in flow) or `wp:anchor` (floating); size `wp:extent cx/cy` in EMU
  (360,000 per cm); alt text `wp:docPr descr/title`; image bytes via `a:blip r:embed` -> rels ->
  `/word/media/...`. Text boxes: `w:txbxContent` (its paragraphs are not body paragraphs).
- Styles: `w:style w:type w:styleId` + `w:name`. **Built-ins are spelled in lowercase in the XML**
  (`heading 1`, `caption`, `toc 1`) while Word displays `Heading 1`, `Caption`, `TOC 1`
  (`display_style()`); `w:customStyle="1"` marks the document's own styles; `w:basedOn`,
  `w:pPr/w:outlineLvl` (0-8 -> heading level 1-9; this is how company heading styles are recognised),
  `w:pPr/w:numPr` (a list style).
- Numbering: `w:abstractNum` (formats per level: `w:numFmt` bullet/decimal/lowerLetter..., `w:lvlText`,
  indents) and `w:num` (an instance pointing at an abstractNum, optional `w:lvlOverride/w:startOverride`).
  A paragraph's list is `w:numPr` (`w:numId`, `w:ilvl`) directly or through its style.
- Comments: `comments.xml` `w:comment w:id w:author w:date`; anchors in the body as
  `w:commentRangeStart/End` + `w:commentReference`. Threads and resolution live in
  `commentsExtended.xml`: `w15:commentEx w15:paraId` (= the comment's last paragraph `w14:paraId`),
  `w15:paraIdParent` (reply-to) and `w15:done="1"` (resolved). Word's `Comments(n)` order equals the order
  of `w:commentReference` in the body (verified, including replies and out-of-order creation).
- Captions: `SEQ <Label>` fields (`fldSimple w:instr=" SEQ Figure \* ARABIC "` or complex fields); the
  label in use ("Figure", "Fig", "Abbildung") is read from the SEQ instruction.

## 7. Writing with Range.InsertXML

`Range.InsertXML(flat_opc_string)` imports a whole fragment in one call. `md_to_word.py` builds the
fragment: `/_rels/.rels`, `/word/document.xml`, `/word/_rels/document.xml.rels`, `/word/styles.xml`,
optional `/word/numbering.xml`, and media parts (`pkg:binaryData`, `pkg:compression="store"`).

Behaviour established by experiment:

- **Merge**: inserting into an empty paragraph, the fragment's last paragraph merges with it - no extra
  empty paragraph is left (verified at the document end and in the middle).
- **Style resolution** (most important):
  - a style referenced by `styleId` that the fragment does not define falls back to **Normal**;
  - if the destination already has the style, the destination's definition wins - a **name-only**
    definition (`<w:style w:styleId=".."><w:name w:val=".."/></w:style>`) is enough, and no duplicate
    is created (custom company styles, tables included);
  - a built-in style the document has **not used yet** is latent: the fragment's definition would be
    used instead of the template's. Fix: instantiate it first (`s = d.Styles(id); s.Priority =
    s.Priority`), see `prepare_styles()`;
  - marking a **built-in** style `w:customStyle="1"` makes Word create a duplicate ("Body Text1").
    Only the document's own custom styles get that flag.
- **Numbering**: `numId`/`abstractNumId` are remapped on import. One `w:num` per markdown list (with
  `startOverride 1` on every level) makes each list restart. A list definition copied from the
  document (`house_style()[kind]["abstract_xml"]`) gives the document's own bullet glyph/format; drop
  its `w:nsid` and `w:tmpl` so Word treats it as a new list instead of continuing an old one.
- **Pictures**: PNG/JPEG/GIF embed via `w:drawing` + a media part. **SVG cannot be embedded through XML**
  without a raster fallback, so SVGs are left as an invisible marker run and replaced afterwards with
  `InlineShapes.AddPicture(svg)`.
- **Captions** need `Fields.Update()` over the inserted range, or every new caption shows "1".
- **Hyperlinks**: relationship with `TargetMode="External"`.
- Fragment styles we define as fallbacks: Heading1-6, ListParagraph, Quote, Caption, Code (custom),
  Inline Code (custom character), Hyperlink, TableGrid.

## 8. Styles and lists

- `Styles(name)` or `Styles(WdBuiltinStyle id)`; `NameLocal` is the displayed (possibly localized) name;
  use ids for built-ins to survive localization (`BUILTIN_STYLES` in the server).
- A new paragraph created by splitting inherits its neighbour's style **and list numbering**: a table
  appended after a numbered list had every cell numbered (3..12) until `plain_paragraphs()` started
  resetting new paragraphs to Normal and calling `ListFormat.RemoveNumbers()`.
- **House style** (`Context.house_style()`): for each role (heading1-6, title, body, bullet, number,
  table, table_text, code, inline_code, quote, caption, figure, caption_label) the style the document
  *uses* wins; then a matching style its template defines but nobody used yet; then Word's built-in.
  `word_apply_styles` uses `CopyStylesFromTemplate` to bring a template's styles into a document.
- Creating a custom bullet in tests: `range.ListFormat.ApplyBulletDefault()`, then
  `ListTemplate.ListLevels(1).NumberFormat = "»"` and re-apply the template (setting `NumberStyle`
  late-bound fails with `0x800A1200`).

## 9. Comments, revisions, fields, pictures, pages

- `Comments.Add(range, text)` returns the comment; `comment.Replies.Add(range, text)` adds a reply
  (replying to a reply: use the thread root, `comment.Ancestor`); `comment.Done` resolves; replies'
  `Ancestor` is the root (`None` for a root - test with `is None`, pywin32 COM objects have no reliable
  truthiness).
- Changing a comment-anchored paragraph's text through `Range.Text` deletes the comment and its replies
  with it (observed when a table cell carrying a comment was rewritten).
- `Revisions.AcceptAll()/RejectAll()`; revision `Type` codes 1 insert, 2 delete, ... (see
  `word_track_changes`).
- `TablesOfContents.Add(range, UseHeadingStyles, Upper, Lower)` then `Update()`; `Fields.Update()`
  refreshes SEQ numbers, cross references, TOC.
- `InlineShapes.AddPicture(path, LinkToFile=False, SaveWithDocument=True, range)`; scale with
  `Width`/`Height` keeping the ratio; `AlternativeText` for alt text. `InlineShapes(k)` order equals the
  order of inline `w:drawing`s in the XML (floating `Shapes` are a separate collection).
- `Range.InsertCaption("Figure", ": title", "", 1)` adds a numbered caption below.
- Page images (`EnhMetaFileBits`) show text, tables, pictures, headers/footers but **not comment
  balloons**; at 96 dpi thin table borders can drop out (rendering artefact, not a document problem).

## 10. Surprising behaviours

| Behaviour | Consequence / handling |
|---|---|
| `Documents.Open` replaces a lone, untouched blank document (like File > Open) | `word_open` reports `replaced_blank_document`; tests open copies while another document is open |
| Documents in **Protected View** are not in `Documents` *(unverified detail: they live in `ProtectedViewWindows`)* | tools say "no document"; the user must click Enable Editing |
| `DeleteAllComments` raised "This command is not available" (no comments present) | delete `Comments(i)` from the end |
| `TrackRevisions` raises while a comment card is selected in the comments pane | read it with `safe()`; status/context still work |
| Word refuses automation while a dialog/menu is open or the scrollbar is dragged | retry 5 s then "Word is busy" |
| A user scrolling/typing makes automation ~1.6x slower (UI thread busy) but never wrong | measured by `test_concurrent_user.py` |
| Background `WM_VSCROLL` to Word's top-level window does nothing (the page is a child window); background PageUp/PageDown keystrokes do scroll and move the caret | how `test_concurrent_user.py` simulates a reader |
| Undo history is shared with the user | `word_undo` right after the user types undoes the user's typing |
| `ScreenUpdating=True` triggers a repaint (8 ms+) | only freeze for bulk edits |
| `ComputeStatistics` forces pagination | `word_context(stats=false)` on huge documents |
| AutoSave on SharePoint/OneDrive uploads every edit | `word_autosave(false)`; the agent never saves unless asked |
| `python -I` (how clients launch the server) leaves the script folder off `sys.path` | the server and tests insert their own folder explicitly before importing local modules |
| f-strings reusing the same quote inside `{...}` only parse on Python 3.12+ | keep code valid for Python 3.10 (the README minimum) |
| MCP Python SDK 2.x renamed FastMCP | `requirements.txt` pins `mcp>=1.10,<2` |

## 11. word-mcp architecture

```
agent --MCP stdio--> word_mcp_server.py --queue--> ComWorker (1 STA thread) --COM--> Word (WWLIB.DLL)
                         |  docx_context.py: parse WordOpenXML -> markdown view, paragraph map,
                         |                   comments, images, house style
                         |  md_to_word.py:   markdown -> flat-OPC fragment for Range.InsertXML
                         |  dump_typelib.ps1: MSWORD.OLB -> word-api.json (word_api_search)
```

Key functions in `word_mcp_server.py`:

| Function | Purpose |
|---|---|
| `ComWorker._call` | runs a closure on the COM thread; busy retry, reconnect, transient `<unknown>` retry, logging |
| `ComWorker.run(fn, edit=True)` | marks agent edits: afterwards `after_agent_edit` re-baselines `word_changes` and bumps `EDITS` |
| `get_context(d, refresh)` | cached `Context(d.WordOpenXML)`; fingerprint = hash(Content.Text), picture/shape/comment/revision counts, `EDITS` |
| `paragraph_map(d)` | **the** paragraph numbering (XML parse == Word); walks Word's paragraphs if counts differ |
| `locate(...)` | number or anchor -> paragraph; exactly-one rule; table/TOC exclusion; `expect`; `verify_paragraph` |
| `verify_paragraph(d, n, m)` | compares Word's paragraph n with the map (alphanumerics only); refuses on mismatch |
| `section_end(d, n, pm)` | where a heading's section ends; stays before a trailing section break |
| `new_paragraph_at(...)` | creates one empty Normal paragraph at end/start/after/before/section_end/cursor; "after" splits at the end of the anchor paragraph (never enters a following table) |
| `neighbours(d, first, last)` | `placed_after` / `placed_before` evidence in insert results |
| `house_for(d, mode, overrides)` | styles for new content |
| `undo_step(app, label, freeze)` | one Ctrl+Z step, optional ScreenUpdating freeze |
| `safe(fn, default)` | property reads Word may refuse |
| `find_table(d, table, near)` / `read_table(d, t)` | a table by number or by unique text in a cell; its cells as a grid (merged cells walked one by one) |

Invariants - keep them when changing code:

1. Paragraph numbers always come from `paragraph_map()`; never index `d.Paragraphs` with a number
   computed from `Content.Text`.
2. Every edit at a paragraph goes through `locate()` (which verifies against Word) or verifies the
   target itself.
3. Anchors resolve to exactly one paragraph or are refused; never fall back to a different paragraph
   that merely contains the words when the text names a table cell or contents line.
4. Destructive edits (delete, rewrite) refuse to remove pictures, comments, tables, section breaks
   unless `allow_content_loss=true`, and report what they removed.
5. Agent edits run with `edit=True` so `word_changes` only reports the user's edits.
6. One COM thread; bulk reads and writes; no per-paragraph loops in hot paths.
7. The server never saves unless a tool explicitly asked to (`word_save`, `word_close(save=true)`).

`docx_context.Context` keeps per paragraph (`para_meta`): `n`, `sid`, `style` (display name), `plain`
text, `in_table`, `row_end`, `heading` level, `list` kind and `abs` (abstractNum), `image`, `hidden`,
`section_break`. `render()` produces the `[pN]` markdown view; `comments()`, `images`, `house_style()`,
`style_catalog()`, `headers_footers()`, `notes()`, `properties()` provide the rest.

## 12. Troubleshooting playbook

| Symptom | Likely cause | Fix / where to look |
|---|---|---|
| Insert landed in the wrong place but returned success | numbering mismatch or a loose anchor | check `placed_after/placed_before`; run `test_anchors.py`; any new numbering source must match `paragraph_map()` |
| "safety check failed" | Word's paragraph differs from the map (document changed, or a parse/numbering bug) | re-read; if it repeats on an unchanged document, compare `paragraph_map(d)` with a walk of `d.Paragraphs` (see `repro` in §13) |
| "anchor ... is ambiguous" / "only appears in ... table" | anchor matched several paragraphs / only table text | longer anchor, `in_tables=true`, or `paragraph`+`expect` |
| A delete removed more than asked | range included pictures/comments/breaks | deletes now refuse; reproduce with the fixture and inspect `removed_content` |
| Headings/tables do not look like the template | style latent, custom flag on a built-in, or wrong role detection | `word_styles`; §7 style rules; `test_styles.py` |
| Each new list continues the previous numbering | copied `w:nsid` kept, or no `startOverride` | `_house_abstract()` / `_numbering()` in `md_to_word.py` |
| Captions all say 1 | fields not updated | `Fields.Update()` on the inserted range / `word_update_fields` |
| A table appears numbered like a list | new paragraphs inherited list numbering | `plain_paragraphs()` must run on new paragraphs |
| "Word is busy" | dialog/menu open, or the user dragging the scrollbar | close it; nothing to fix in code |
| "Object has been deleted" in tests | two suites driving the same Word | run suites one at a time (`tests_common.setup()` locks) |
| `UnicodeEncodeError` printing | cp1252 console | `tests_common.setup()` switches stdout to UTF-8 |
| `<unknown>.Name` | document mid-save/sync | retried automatically; check `WORD_MCP_LOG` if persistent |
| `word_read` / `paragraph_map` is 2+ paragraphs short, "safety check failed" on a real document with a cover page | table cells or rows wrapped in a content control (`w:sdt`, e.g. a cover table bound to document properties) were skipped by the parse | `Context._unwrapped()` looks through `w:sdt` for `w:tr`/`w:tc`; pinned by `test_sdt_table.py` (offline). Compare `len(Context(d.WordOpenXML).para_meta)` with `d.Paragraphs.Count` |
| `word_status` fails on one property | Word refuses it in the current UI state | wrap the read in `safe()` |
| `No module named 'mcp.server.fastmcp'` | MCP SDK 2.x | `pip install "mcp>=1.10,<2"` |
| `word_page_image` returns the wrong page | client cached an old schema (`page` vs `pages`) | both are accepted; restart the client |
| A tool works on a hidden test document but not on the user's | view/UI state (Protected View, comment pane, Print Layout) | read state with `safe()`, restore any view you change |

## 13. How to fix and verify a bug

1. **Never debug on the user's document.** Build a disposable one. `tests_fixtures.build_integration_doc`
   reproduces a real integration-overview layout: title, TOC, references and overview tables repeating
   provider names, a layout table with a nested table, two section breaks with a landscape section,
   headings H1-H3, two images with SEQ captions, comments, a manual page break, hidden text, a content
   control, a tracked deletion.
2. **Reproduce by comparing sources of truth**, e.g. for numbering:

   ```python
   real, p = [], d.Paragraphs(1)
   while p is not None:              # Word's own paragraphs (slow, exact)
       real.append(p.Range.Text); p = p.Next()
   ctx = Context(d.WordOpenXML)      # what word-mcp believes
   # compare len(real), ctx.para, len(ctx.para_meta), and texts by alphanumerics
   ```

3. **Assert the real document state**, not the tool's own response: read Word directly over COM in the
   test (`Range.Information(12)` for "in a table", `Information(2)` for the section, `Sections.Count`,
   `InlineShapes.Count`, `Comments.Count`).
4. **Add a regression check** to the matching suite (`test_anchors.py` for placement/numbering,
   `test_styles.py` for styling, `test_tables.py` for tables, `test_context.py` for reading/comments/images, `test_v2.py` for
   markdown/changes/batch, `test_word_mcp.py` for the basic tools, `test_concurrent_user.py` for a user
   working at the same time).
5. **Run every suite, one at a time**, all must end with `ALL CHECKS PASSED`:

   ```powershell
   python -I test_anchors.py; python -I test_tables.py; python -I test_word_mcp.py; python -I test_context.py
   python -I test_v2.py; python -I test_styles.py; python -I test_concurrent_user.py
   ```

6. Restart the MCP client to load the new server (clients cache tool schemas: keep old parameter
   names working when renaming one).

## 14. Re-deriving this

- **Binaries**: REA `inspect_pe_resources` on `WINWORD.EXE` and `MSWORD.OLB` (resources), a PE
  import/export read of `WINWORD.EXE` and `WWLIB.DLL`. `WWLIB.DLL` is too large for full Ghidra
  analysis within REA's 330 s limit; nothing above needs it.
- **API**: `powershell -File dump_typelib.ps1` (registry lookup of the type library, `LoadTypeLibEx`,
  every type, member, parameter and enum constant).
- **Behaviour**: small probes against a hidden document (`Documents.Add(Visible=False)`), always closed
  with `Close(0)` in a `finally`; measure with `time.perf_counter()`; look at pages with
  `EnhMetaFileBits` rather than trusting text alone (that is how the numbered-table bug was found).
