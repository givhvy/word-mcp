# Issues observed while using word-mcp

Observed on Windows with Word 16.0.20326, Python 3.13, and the OpenCode `word-new` connection. The checkout is `d35b0d4` at `D:\wordfolder\word-mcp`. Tests below used disposable documents unless stated otherwise. The live document was a SharePoint `.docx` open in Word desktop, with layout tables, an overview table, images, captions, and comments. This is a report of observations, not a claim that every failure is reproducible on every Word installation.

## 1. Markdown insertion can target the wrong part of a document

**Severity:** High. It can modify a live document in an unrelated place while returning success.

In the SharePoint integration overview, the body has a `VeloViewer` Heading 2, but `VeloViewer` also appears in the References table and in the overview table. I tried to add a linked sentence at the end of that section:

```text
word_insert_markdown(
  markdown="For further details on VeloViewer, see [D0180 - Integration Analysis - VeloViewer](https://.../D0180%20-%20Integration%20Analysis%20-%20VeloViewer.docx?web=1).",
  position="section_end", anchor="VeloViewer", document=<open SharePoint URL>)
```

The tool returned success, but `word_context` showed the sentence after **Conclusion**, at the end of the document, not after the VeloViewer section. Another attempt with `position="before", anchor="Google Street View"` inserted it **inside the overview table**, before its Google Street View row. A third attempt with `position="after", anchor="Figure 4. VeloViewer integration flow."` returned success but `word_context` showed it between `How it is called / request body` and `Data received` in VeloViewer, not after Figure 4. Before that third attempt, the same tool with `paragraph=202, expect="Figure 4. VeloViewer integration flow."` refused the edit because it claimed paragraph 202 was Google Street View protocol, though `word_read` had just reported paragraph 202 as the caption.

**Expected:** The tool should land after the uniquely identified Heading 2 or caption, or refuse the edit if the anchor is ambiguous or the paragraph indexes disagree. It must not silently choose a table cell or another section.

**Likely area to inspect:** `locate()`, `paragraph_texts()`, `section_end()`, `new_paragraph_at()`, and the position used by `word_insert_markdown()` in `word_mcp_server.py`. `locate()` takes the first normalized exact/prefix/substring match without checking its style or table membership. `section_end()` uses heading numbers from `Context` while insertion uses `d.Paragraphs(n)`; these numbering systems appeared to disagree in this document. The tool's returned `first_paragraph`/`last_paragraph` did not identify the actual paragraph shown by `word_context`.

**Regression test:** Create a Word document with a references table, an overview table repeating a provider's name, a Heading 2 using that name, several image/caption paragraphs, and another Heading 2. Insert at the provider's `section_end`, before the next heading, and after a caption. Assert the exact resulting OOXML/body order, not just the tool response. Test ambiguous anchors and `paragraph` plus `expect` too.

I used `word_undo` to back out the misplaced attempts. In the course of recovery, paragraph and image counts briefly changed; the last `word_context` read after the undos showed the original 17 images and 9 comment threads, and the document reported `saved`. The user's request was **not** completed or the comment resolved.

## 2. Paragraph deletion on an image-rich document is not predictably one paragraph

**Severity:** High. A narrow cleanup command coincided with other content disappearing.

While removing the misplaced VeloViewer link above, I called `word_delete_paragraphs(anchor="For further details, see D0180 - Integration Analysis - VeloViewer.", expect=<same text>)` several times as the sentence still appeared in `word_read` after calls that returned `{"deleted": 1}`. During those attempts the document dropped from 17 to 16 images and from 9 to 8 comment threads according to `word_context`. `word_undo` restored the original counts. It is not established whether deletion alone, a shifted paragraph index, or an earlier insertion caused the loss; the counts make this worth a targeted regression test.

**Expected:** Deleting a matching plain paragraph either removes only that paragraph or refuses the operation if its location is unsafe. A repeat call should report that the anchor is gone, not appear to delete unrelated content.

**Likely area to inspect:** `word_delete_paragraphs()` resolves an anchor through `locate()` and then deletes `d.Paragraphs(first).Range` to `d.Paragraphs(last).Range`. Check the mapping between `paragraph_texts(d)` and Word's `Paragraphs` collection in documents with tables, images, and comments.

## 3. Fresh install selects an incompatible MCP major version

**Severity:** Setup blocker. Initially observed on the original checkout; still applicable to `d35b0d4`'s upstream requirements.

The upstream `requirements.txt` says `mcp>=1.10`. On a fresh Python 3.13 virtual environment, pip selected MCP 2.3.0. The server then failed immediately at `from mcp.server.fastmcp import FastMCP, Image` with:

```text
ModuleNotFoundError: No module named 'mcp.server.fastmcp'. This is mcp 2.x, where FastMCP was renamed to MCPServer
```

Local workaround: change the requirement to `mcp>=1.10,<2`, reinstall, and launch again. The local checkout still has that uncommitted one-line compatibility pin. Upstream should either declare `<2` or migrate the imports and API to MCP 2.

## 4. `word_status` fails when a Word comment card is selected

**Severity:** Medium. Observed while working on the open SharePoint document.

With a comment card selected in Word's comments pane, `word_status()` returned:

```text
Word: The TrackRevisions method or property is not available because comment card is selected in pane. (0x80020009)
```

Other calls such as `word_read_tables()` and `word_list_comments()` still worked. `word_status()` reads `d.TrackRevisions` while building the active-document summary. The new `word_context()` also reads `d.TrackRevisions`; it may need the same protection, though that particular failure was observed on `word_status`, not `word_context`.

**Expected:** Status still reports the open documents, with `track_changes` unavailable or omitted when Word refuses that property. Reading metadata should not fail in full because a comment card has focus.

## 5. Tests in the current update fail in this environment

**Severity:** Medium for release confidence. Observed at `d35b0d4` using `.venv\Scripts\python.exe -X utf8 -I ...`.

- `test_word_mcp.py` completed 48 tool calls but ended with `FAIL user's documents untouched, scratch docs closed`. This was after other scratch tests had been run against the same Word session, so an orphaned scratch document is a possible explanation. The test does not identify which document differed in its final assertion.
- `test_v2.py` completed its workflows but ended with `FAIL heading uses the template font (Aptos Display)` in the scratch document. Investigate whether `word_insert_markdown()` failed to follow the template font or whether the test assumes a font not present in this Word/template setup.
- Running `test_v2.py` without Python UTF-8 mode stopped early with `UnicodeEncodeError` while printing a Vietnamese character to the Windows cp1252 terminal. `-X utf8` let it continue. A previous `test_context.py` run had the same terminal-output issue; it passed with `-X utf8`.
- A first parallel test run also produced `Word: Object has been deleted. (0x80020009)` during `word_format(find="budget", bold=True, highlight="yellow")` in `test_word_mcp.py`. A later sequential rerun passed that step. This points to either shared Word COM state under concurrent test runs or an intermittent stale object, not a proven deterministic `word_format` failure.

**Suggested checks:** Make test cleanup identify leftover scratch document names; avoid concurrent suites that automate the same Word instance; set UTF-8 for test stdout; and run the heading-font assertion against an explicitly chosen template. Keep separate tests for concurrent *user navigation* rather than concurrently running two full automation suites.

## 6. Comments and saves can be awkward during live editing

**Severity:** Low; needs reproduction before treating it as a code defect.

Earlier, after resolving a comment in the SharePoint file, an immediate `word_save()` once returned `Error executing tool word_save: <unknown>.Name`; a second `word_save()` succeeded and the comment remained resolved. This was a transient result with no captured COM trace. Log the underlying exception and which Word object was being accessed rather than returning `<unknown>.Name` if it happens again.

## Notes for the maintainer

- The Word connection itself works: the new server exposed 44 tools, read the open SharePoint file and its unsaved changes, edited table cells, and resolved comments. This report focuses on failures and near misses.
- The live SharePoint file has user edits outside this investigation. Do not use it as a destructive test fixture. Reproduce with a disposable `.docx` containing the same structural features.
- I have not benchmarked old versus new MCP end to end on the same document. The published speed claims are not a measured regression here.
