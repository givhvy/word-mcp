# Reverse engineering Microsoft Word: the full analysis record

This is the complete record of how word-mcp's knowledge of Word was obtained: what was analysed, with which
tool, what came out (with hashes so results can be checked), what it means, and **how to reproduce every
step on a PC that does not have REA**. Together with [WORD_INTERNALS.md](WORD_INTERNALS.md) (what was
learned) and [WORD_API_REFERENCE.md](WORD_API_REFERENCE.md) (the generated API reference), it is everything a
maintainer needs to understand and fix word-mcp.

**About "source code".** Word is closed source. Reverse engineering here means analysing Word's *shipped
binaries* and its *published automation interface*; it does not produce, and this repository does not
contain, Microsoft source code or decompiled Microsoft code (that would also not be ours to redistribute).
Nothing in word-mcp needs it: the server only uses Word's documented automation API (described by Word's
own type library, which every Word install carries) and the document's own XML (WordOpenXML). Those two
are fully covered below, and the source code that matters for fixing word-mcp - the server itself - is in
this repository.

Analysed install: Microsoft 365, `C:\Program Files\Microsoft Office\root\Office16\`, WINWORD.EXE
16.0.14332.20447, WWLIB.DLL 16.0.14332.20438, MSWORD.OLB 16.0.10001.10000, Windows 11 x64.

---

## 1. Tools

| Tool | Used for | Needed to reproduce? |
|---|---|---|
| REA (Reverse Engineer Anything, `rea-agents` 6.4.0, Ghidra 12.1.4 provider) | PE resource inspection of `WINWORD.EXE` and `MSWORD.OLB` | **No** - `tools/pe_inspect.py` reproduces every binary result with the Python standard library |
| `dump_typelib.ps1` (in this repo) | decode `MSWORD.OLB` with Windows' own `LoadTypeLibEx` into `word-api.json` | ships with word-mcp |
| `tools/api_reference.py` | turn `word-api.json` into `docs/WORD_API_REFERENCE.md` | ships with word-mcp |
| Python + pywin32 probes | measure behaviour through COM on hidden scratch documents | ships with the tests |
| `Document.WordOpenXML` | read the full document structure as XML | built into Word |

## 2. Step 1 - find the binaries

Word's process is `WINWORD.EXE`; next to it in `root\Office16`:

| File | Bytes | SHA-256 |
|---|---|---|
| `WINWORD.EXE` | 1,637,280 | `28e2731d1fc2cf95f341969e5cebe7b09d453cfd4562986d83f18888178dba8d` |
| `WWLIB.DLL` | 45,763,960 | `f314e19ab901365615731cdc9d0c8ab5bfda172899855fa203c19f2814826e9f` |
| `MSWORD.OLB` | 952,112 | `bf8c4d0d2a6656c70c0894332da48c9d42b147972ca256ae977e8c0c960b94c5` |
| `Wordcnvpxy.cnv` | 50,464 | (not analysed) |

Hashes differ on other builds; the structure does not.

## 3. Step 2 - WINWORD.EXE: a launcher

**REA** `inspect_pe_resources(WINWORD.EXE)`: coverage complete, 490 directory entries, **242 resources**:

| Resource type | Count | Bytes |
|---|---|---|
| ICON | 225 | 1,566,026 |
| GROUP_ICON | 13 | 3,228 |
| MANIFEST | 1 | 4,496 |
| VERSION | 1 | 1,020 |
| EDPENLIGHTENEDAPPINFOID / EDPPERMISSIVEAPPINFOID | 1 + 1 | 2 + 2 (Windows Information Protection markers) |

No dialogs, menus or strings - Word's UI is not here.

**PE structure** (`tools/pe_inspect.py`): PE32+ x64, sections `.text` (15 KB of code), `.rdata`, `.data`,
`.pdata`, `.didat` (delay-load imports), **`.c2r`** (Click-to-Run), `.rsrc` (1.59 MB - the icons), `.reloc`.
Exports: `DllGetLCID`. Imports: `AppVIsvSubsystems64.dll` (App-V / Click-to-Run virtualization), ADVAPI32
registry calls, KERNEL32 (mutex, thread, events), and the C/C++ runtime. **No import of WWLIB.DLL**, yet the
file contains the UTF-16 string `wwlib` and the ASCII strings `LoadLibrary`, `GetProcAddress` and `FMain`.

**Conclusion:** `WINWORD.EXE` is a 15 KB-of-code launcher that loads `WWLIB.DLL` at run time and calls its
`FMain` export. Word itself lives in `WWLIB.DLL`.

## 4. Step 3 - WWLIB.DLL: the engine

(`tools/pe_inspect.py`; REA was not run to completion on this file - see step 6.)

- PE32+ x64; `.text` **28.7 MB of code**, `.rdata` 6.3 MB, `.data` 1.5 MB, `.rsrc` 7.4 MB, `.didat`
  (delay-load imports), `_RDATA`.
- **Exports (19):** `FMain` (Word's main, called by the launcher), `DllMain`, **`DllGetClassObject`** (the COM
  class factory - this is how `Word.Application` objects are created for automation), `DllCanUnloadNow`,
  `DllGetLCID`, plus 14 C++ exports in a `WordMailReact` namespace (Word as Outlook's mail editor:
  `CreateReadingPaneDocument`, `InsertPartsIntoDocument`, `StartSave`, ...).
- **Imports (22 DLLs):** `oart.dll` by ordinal (Office Art: shapes, pictures, charts), GDI32 and gdiplus
  (rendering; `CloseEnhMetaFile` - the EMF page rendering that `word_page_image` uses comes from here),
  ole32/OLEAUT32 (OLE embedding, COM marshaling, VARIANT/BSTR), WINMM, ADVAPI32 (registry, event log), the
  C/C++ runtime.
- **Resources (1,009):** 121 `UIFILE` (ribbon/UI definitions), 463 cursors, 225 icons, 16 bitmaps, a
  `WEVT_TEMPLATE` (event-log manifest), a `MESSAGETABLE`, 2 `RCDATA`, and 73 resources of private type 998.
- Strings: the automation member names (`InsertXML`, `WordOpenXML`, ...) do **not** appear in the DLL as
  ASCII or UTF-16 text; neither do OOXML element names such as `w:sectPr`.

**Conclusion:** all document logic, layout and file formats are in `WWLIB.DLL`, but nothing in it is
addressed by name from outside except through COM - whose names live in the type library (next step).

## 5. Step 4 - MSWORD.OLB: the automation contract

**REA** `inspect_pe_resources(MSWORD.OLB)`: a PE32+ image (machine `0x8664`) whose resource directory sits
at RVA 8192 / file offset 1,536 and holds exactly two resources:

| Resource | Lang | File offset | Bytes | SHA-256 |
|---|---|---|---|---|
| `TYPELIB` #1 | 1033 | 2,736 | **931,600** | `41c6feaa56d473d31cd14844757ed807cc36562ddb9bc06edbfd085d91032739` |
| `VERSION` #1 | 1033 | 1,712 | 1,024 | `c2c6626c82434629fae93da0938e0df4d4e4614ac4cca3f77ee4abecf3e6eb34` |

`tools/pe_inspect.py` finds the same offsets, sizes and hashes (and `--extract-typelib out.tlb` writes the
payload to a file).

Registry: `HKCR\TypeLib\{00020905-0000-0000-C000-000000000046}\8.7\0\win64` points at this file -
type library GUID `{00020905-0000-0000-C000-000000000046}`, version **8.7**.

**Decoding** (`dump_typelib.ps1`): `LoadTypeLibEx(path, REGKIND_NONE)` then, for every type,
`ITypeInfo.GetTypeAttr` / `GetFuncDesc` / `GetVarDesc` / `GetNames` / `GetDocumentation`, resolving
pointer, SAFEARRAY and user-defined parameter types. Output `word-api.json`: a list of
`{name, kind, guid, doc, members: [{name, kind (method/get/put/putref/const), dispid, returns, hidden,
params: [{name, type, optional}]}]}`.

Result for this build:

| | Count |
|---|---|
| Types | 767 |
| Dispatch interfaces | 374 |
| Enums | 385 (3,754 constants) |
| Coclasses | 7 |
| Plain interfaces | 1 |
| Visible members on dispatch interfaces | 10,445 |
| Methods among them | 4,072 |

Biggest interfaces: `_Document` 386 members, `_Application` 231, `Range` 214, `Selection` 208,
`Paragraphs` 107, `Find` 80, `Table` 70, `ContentControl` 64. Members flagged *hidden* are legacy but still
callable (word-mcp uses `Comment.Author` and `Comment.Delete`, both hidden).

`docs/WORD_API_REFERENCE.md` is generated from this: every member, with signatures, of the 60 interfaces
and 29 enums word-mcp works with, ticking the members the server calls.

## 6. Step 5 - why not decompile WWLIB.DLL

In the same environment, REA's Ghidra provider was run on a smaller native DLL - FL Studio's 22 MB
`FLEngine_x64.dll`, an unrelated program analysed earlier - with REA's 330-second provider timeout:
`search_strings` timed out, and an `rea inspect --snapshot` run failed with `provider_timeout` (operation
`list_documents`, 330,000 ms) after Ghidra worked for about 9 minutes. `WWLIB.DLL` is twice that size
(45.8 MB, 28.7 MB of it code), so it was not attempted. A full analysis would need a detached,
resumable Ghidra job taking hours, and would yield decompiled internals that are neither redistributable
nor needed: every capability word-mcp uses is exposed - by design - through the type library and
WordOpenXML. So the analysis stopped at the binary-structure level for WWLIB.DLL and went deep on the
two public surfaces instead.

## 7. Step 6 - the document surface (WordOpenXML)

`Document.WordOpenXML` returns the live document (unsaved edits included) as one flat-OPC XML package.
Parts, element names and behaviours used by word-mcp are catalogued in
[WORD_INTERNALS.md section 6](WORD_INTERNALS.md#6-wordopenxml). Method: dump the XML of small test
documents built one feature at a time (a heading, a list, a table, a comment with replies, a section
break, a caption...) and read which elements appear; then check every conclusion against Word's own
object model (for example, paragraph counts against `Document.Paragraphs.Count`).

## 8. Step 7 - behaviour by experiment

Everything the binaries and the type library cannot say - what Word *does* - was measured with small
probes on hidden documents (`Documents.Add(Visible=False)`, always closed with `Close(0)`):

- costs of COM calls, bulk reads/writes, page rendering (WORD_INTERNALS.md section 3);
- the paragraph-numbering rules and the Content.Text section-break mismatch (section 4);
- how `Range.InsertXML` resolves styles, lists and pictures (section 7);
- comment threading and ordering, field updates, page rendering limits (section 9);
- surprising behaviours and error codes (sections 3 and 10).

Each finding that matters is pinned by a test in the repository (`test_*.py`), so a Word update that
changes a behaviour shows up as a failing check.

## 9. Reproduce everything without REA

```powershell
cd word-mcp
python tools\pe_inspect.py                              # steps 1-4: hashes, sections, resources, imports, exports
python tools\pe_inspect.py --extract-typelib word.tlb   # the TYPELIB payload, hash-checked
powershell -ExecutionPolicy Bypass -File dump_typelib.ps1   # step 4: word-api.json from your Word
python tools\api_reference.py                           # docs/WORD_API_REFERENCE.md for your build
python -I test_anchors.py                               # steps 6-7 pinned as tests (run suites one at a time)
```

Compare your hashes and counts with the tables above to see whether your Word build differs; if the type
library differs, regenerate `WORD_API_REFERENCE.md` and look up any member a failing tool uses.

## 10. Using this to fix word-mcp

1. Read the error: Word's own message is in the tool error (`Word: ... (0x80020009)`); set
   `WORD_MCP_LOG` for the full trace. Match it in [WORD_INTERNALS.md section 12](WORD_INTERNALS.md#12-troubleshooting-playbook).
2. Find the Word API the failing code calls in `word_mcp_server.py`, then its exact signature and
   alternatives in [WORD_API_REFERENCE.md](WORD_API_REFERENCE.md) (or `word_api_search`).
3. Reproduce on a disposable document (`tests_fixtures.build_integration_doc`), never the user's.
4. Fix, keep the invariants in [WORD_INTERNALS.md section 11](WORD_INTERNALS.md#11-word-mcp-architecture),
   add a check to the matching test suite, run all suites one at a time.
