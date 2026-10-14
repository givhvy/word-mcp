# word-mcp

An MCP server that lets an AI agent work inside the **Microsoft Word that is open on your Windows PC**:
read the whole document with its structure, see its images and pages, write and format text, build
tables, insert pictures, and hold comment conversations (add, reply, resolve), live, while you watch.

It talks to Word through Word's own COM automation API, so everything happens in the real Word window,
and every agent edit is a single Ctrl+Z step.

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
```

Each test creates its own scratch documents in your open Word, checks every tool, and closes only what
it created. Both should end with `ALL CHECKS PASSED`.

## Guide for agents

1. `word_status`: which documents are open and which is active.
2. `word_context`: read the document before editing. It returns everything in one call:
   - stats, cursor position, properties, header and footer
   - the body as markdown, every line prefixed with Word's paragraph number `[pN]`. It shows headings
     (`#`), lists, tables, bold and italic, links, tracked changes as `{-old-}{+new+}`, and images as
     `![image n: alt text, WxH cm]`
   - every comment thread with its replies and resolved state, plus footnotes
   - `include_images=true` attaches the images. Long documents continue with `start_paragraph`.
3. Edit by paragraph number (`word_insert`, `word_set_paragraph`, `word_format`, `word_delete_paragraphs`)
   or by text (`word_find`, `word_replace`, `word_format(find=...)`).
4. Paragraph numbers shift after inserting or deleting. Read again (`word_context` or `word_read`)
   before the next edit that uses a number. Comment numbers `n` shift the same way after adding or
   deleting comments.
5. Check the layout with `word_page_image(page)`, which shows the page exactly as Word renders it.
   Text output alone can miss layout problems.
6. New paragraphs are Normal unless you pass `style` ("Heading 1", "Title", "List Bullet",
   "List Number", ...). They never inherit a neighbouring heading or list.
7. Every tool takes an optional `document` (name or full path); the default is the active document.
8. Anything without a dedicated tool: `word_api_search` to look up Word's object model, then `word_com`
   to get, set or call it. Example: path `ActiveDocument.Paragraphs(3).Range.Font.Size`.

## Tools (35)

| Area | Tools |
|---|---|
| Context | `word_context` (whole document in one call), `word_status` |
| See | `word_page_image` (a rendered page), `word_get_image` (an embedded image), `word_list_images` |
| Documents | `word_new_document`, `word_open`, `word_save` (docx, pdf, rtf, txt, html, odt, ...), `word_close`, `word_activate` |
| Read | `word_get_text`, `word_read` (paragraphs, optional styles), `word_find`, `word_read_tables` |
| Edit | `word_insert` (start, end, before or after a paragraph, cursor; optional style), `word_set_paragraph`, `word_delete_paragraphs`, `word_replace`, `word_insert_table`, `word_set_cell`, `word_insert_image`, `word_undo` |
| Format | `word_format` (bold, italic, underline, strikethrough, font, size, color, highlight, style, alignment, spacing; targets a paragraph range, every match of a text, a character range or the selection), `word_page_setup`, `word_header_footer` |
| Comments | `word_add_comment`, `word_reply_comment`, `word_list_comments` (threads), `word_resolve_comment`, `word_edit_comment`, `word_delete_comment` |
| Review | `word_track_changes` (on, off, list, accept all, reject all) |
| Navigate | `word_select` (moves Word's cursor and scrolls to it) |
| Escape hatch | `word_com`, `word_api_search` |

## How it works

- `word_mcp_server.py` (FastMCP, stdio). One dedicated COM thread owns every Word object. It attaches
  to the running Word, retries while Word is busy, and reconnects if Word restarts.
- `docx_context.py` reads `Document.WordOpenXML`, the whole document package in a single COM call,
  and parses it. That is how `word_context`, the comment threads and the images stay fast on
  large documents: about 20-40 ms for the read instead of one round trip per paragraph.
- Bulk edits use Word's own engine: Find/Replace with formatting for "format every match", and
  text-to-table conversion for tables. Each is one call instead of one per item.
- `dump_typelib.ps1` decodes Word's type library (`MSWORD.OLB`, found through the registry) into
  `word-api.json`. `word_api_search` builds it automatically on first use (about 2 s).

Typical speed: simple reads about 10 ms. Edits take 35-150 ms, depending on how busy Word is
(layout and spell-check of visible documents). Whole-document context 0.15-0.2 s, page image about
0.7 s, replace or format 500 matches 0.15-0.4 s.

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
- Works on the main document body. Headers and footers can be set, and are shown read-only in
  `word_context`. Text boxes are shown in context but are not edited by the dedicated tools; use `word_com`.
- Page images do not show comment balloons. Read comments with `word_list_comments`.
