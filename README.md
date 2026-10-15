# word-mcp

An MCP server that lets an AI agent work inside the **Microsoft Word that is open on your Windows PC**,
on local files and on SharePoint / OneDrive documents opened in Word desktop. The agent types straight
into your open document while you watch: no closing, no reopening, no saving behind your back.

Built for writing technical documents with an agent:
- the agent reads your codebase and writes whole sections (headings, lists, code blocks, tables,
  screenshots and diagrams with numbered captions) in **one call** with `word_insert_markdown`;
- it reads the **whole live document**, including what you typed yourself and have not saved, with
  its images and comments (`word_context`), and sees pages exactly as Word renders them;
- it notices **your edits** since it last looked (`word_changes`) and targets places by heading text,
  so your own writing and its writing can be mixed safely;
- saving stays **your decision** (Ctrl+S): the agent never saves unless you ask.

It uses Word's own COM automation API, so everything happens in the real Word window, and every
agent edit is a single Ctrl+Z step.

## Requirements

- Windows 10 or 11
- Microsoft Word desktop (Microsoft 365, 2021, 2019 or 2016). Word on the web and Word for Mac do not work.
- Python 3.10 or newer

## Install

```powershell
git clone https://github.com/givhvy/word-mcp.git
cd word-mcp
python -m pip install -r requirements.txt
```

Check Python can reach Word (prints its version, changes nothing):

```powershell
python -c "import win32com.client as c; print('Word', c.Dispatch('Word.Application').Version)"
```

### Register with your agent

Use the **full path** to `word_mcp_server.py`. The `-I` flag runs Python isolated from the current directory.

**Claude Code**

```powershell
claude mcp add word --scope user -- python -I C:\path\to\word-mcp\word_mcp_server.py
```

Restart Claude Code. The tools appear as `mcp__word__word_context`, `mcp__word__word_insert`, and so on.

**Claude Desktop, Cursor, Windsurf, Cline, or any MCP client using JSON config**

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["-I", "C:\\path\\to\\word-mcp\\word_mcp_server.py"]
    }
  }
}
```

**Codex CLI** (`~/.codex/config.toml`)

```toml
[mcp_servers.word]
command = "python"
args = ["-I", "C:\\path\\to\\word-mcp\\word_mcp_server.py"]
```

If `python` is not on PATH for your client, use the full path to `python.exe`.

### Verify (optional)

```powershell
python -I test_word_mcp.py
python -I test_context.py
python -I test_v2.py
```

Each test creates its own scratch documents in your open Word, checks every tool, and closes only what
it created. Both should end with `ALL CHECKS PASSED`.

## Guide for agents

The server sends these rules to the agent on connect; they are repeated here.

1. **Read first**: `word_context` returns the whole live document in one call: stats, cursor, save state,
   header/footer, the body as markdown with Word paragraph numbers `[pN]` (headings `#`, lists, tables,
   bold/italic, links, code, tracked changes `{-old-}{+new+}`, images `![image n: alt, WxH cm]`), every
   comment thread with replies, and footnotes. `include_images=true` attaches the images;
   `word_page_image(pages="1-3")` shows rendered pages.
2. **Write sections in one call** with `word_insert_markdown`:

   ````markdown
   ## Data flow

   Requests go **client -> API -> DB**. See `api/routes.py`.

   1. Auth checks the token
      1. Validate signature
   2. Handler runs

   ```python
   def handler(req): ...
   ```

   | Layer | Tech |
   |---|---|
   | API | FastAPI |

   ![Architecture](C:/shots/arch.png "System architecture"){width=14cm}
   ````

   Images can be a local path (screenshots, draw.io exports), an https URL or a data: URI; png, jpg,
   gif, svg, emf. The quoted title becomes a numbered "Figure N:" caption. Images are fitted to the
   text width. Rewrite a section with `word_replace_section(anchor="Data flow", markdown=...)`.
3. **Target by text, not just numbers**: `position="section_end", anchor="Components"` appends to a
   section; `position="after", anchor="..."` inserts after any paragraph. When editing by number,
   pass `expect` (the start of the paragraph's text): the edit is refused if the user changed it.
4. **Before the next edit**, `word_changes` lists what the user changed since your last read (added,
   removed and changed paragraphs with current numbers, image/comment counts). It is cheap; the
   agent's own edits are not reported.
5. **Never save** unless the user asks. `word_status`/`word_context` show the save and AutoSave state;
   `word_autosave(false)` turns AutoSave off for a SharePoint/OneDrive document.
6. Comments: `word_add_comment`, `word_reply_comment`, `word_resolve_comment`, `word_list_comments`.
7. Several small steps at once: `word_batch`. Anything without a tool: `word_api_search` + `word_com`.

## Tools (42)

| Area | Tools |
|---|---|
| Read | `word_context` (whole document), `word_changes` (what the user changed), `word_status`, `word_read`, `word_get_text`, `word_find`, `word_read_tables` |
| See | `word_page_image` (rendered pages), `word_get_image`, `word_list_images` |
| Write | `word_insert_markdown` (whole sections), `word_replace_section`, `word_insert` (plain text), `word_set_paragraph`, `word_delete_paragraphs`, `word_replace`, `word_insert_table`, `word_set_cell`, `word_undo` |
| Images | `word_insert_image` (path, base64 or URL; caption; replace an image in place) |
| Format | `word_format`, `word_page_setup`, `word_header_footer`, `word_insert_toc`, `word_update_fields` |
| Comments | `word_add_comment`, `word_reply_comment`, `word_list_comments`, `word_resolve_comment`, `word_edit_comment`, `word_delete_comment` |
| Review | `word_track_changes` |
| Documents | `word_new_document`, `word_open` (path or SharePoint/OneDrive URL), `word_save`, `word_close`, `word_activate`, `word_autosave`, `word_select` |
| Other | `word_batch`, `word_com`, `word_api_search` |

## How it works

- `word_mcp_server.py` (FastMCP, stdio). One dedicated COM thread owns every Word object; it attaches to
  the running Word, retries while Word is busy, and reconnects if Word restarts.
- `md_to_word.py` turns markdown into a WordprocessingML fragment that Word imports with
  `Range.InsertXML`: a whole section, images included, in one call, using the document's own styles.
- `docx_context.py` reads `Document.WordOpenXML` (the whole document in one call) and turns it into
  the markdown view, comment threads and image list. The parse is cached and reused while the document
  is unchanged; images are decoded only when asked for.
- `word_changes` keeps a per-document snapshot of what the agent last saw, refreshed after the agent's
  own edits, and diffs the live paragraphs against it.
- `dump_typelib.ps1` decodes Word's type library (`MSWORD.OLB`, found through the registry) into
  `word-api.json` for `word_api_search`; built automatically on first use (about 2 s).

Measured on Word 16 (Microsoft 365), typical desktop:

| Operation | Time |
|---|---|
| Write a section (headings, nested lists, code block, table, 2 images with captions) | 0.3-0.75 s, one call |
| Write 40 sections with 40 full-width screenshots | 1.0 s, one call |
| Read a 360-paragraph document with 40 images (`word_context`) | 0.24 s first read, 0.1 s when unchanged |
| `word_changes` on that document | 70 ms |
| Simple tool calls | 10-150 ms depending on how busy Word is |
| Page image | 0.2-0.7 s |

The largest cost in practice is the number of agent turns, not Word: prefer `word_insert_markdown`
and `word_batch` over many small calls.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `Word is busy: close any open dialog` | A dialog (Save As, Options, a message box) is open in Word. Close it and retry. |
| `No document is open in Word` | Open or create a document, or call `word_new_document`. |
| `ModuleNotFoundError: win32com` | `python -m pip install pywin32`. If it still fails: `python -m pywin32_postinstall -install` (as administrator). |
| A tool times out | `WORD_MCP_TIMEOUT` (seconds, default 120) sets the limit. A long PDF export of a big document can take a while. |
| `word_open` closed an empty document | Word's own File > Open replaces a lone, untouched blank document. The tool reports it as `replaced_blank_document`. |
| Clients start the wrong Python | Put the full path to `python.exe` in the client config. |

## Limitations

- Windows and Word desktop only, because it uses COM automation.
- Works on the main document body. Headers and footers can be set, and are shown in `word_context`.
  Text boxes are shown in context but are not edited by the dedicated tools; use `word_com`.
- `word_changes` and the context cache see text, image, comment and tracked-change edits; a user's
  formatting-only change is picked up with `word_context(refresh=true)`.
- SharePoint / OneDrive: open the document in Word desktop (not in the browser). `word_open` takes the
  file's direct URL; sharing links (`/:w:/...`) may not open through COM.
- Page images do not show comment balloons. Read comments with `word_list_comments`.
