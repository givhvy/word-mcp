# Agents working on word-mcp

Everything needed to understand and fix word-mcp is in this repository - no REA or other reverse
engineering tool is required:

- [docs/REA_ANALYSIS.md](docs/REA_ANALYSIS.md): how Word was reverse engineered (binaries, hashes, type
  library) and how to reproduce it with `tools/pe_inspect.py` and `dump_typelib.ps1`;
- [docs/WORD_API_REFERENCE.md](docs/WORD_API_REFERENCE.md): signatures of every Word API member the server
  uses (regenerate for your Word build with `python tools/api_reference.py`; search the rest with the
  `word_api_search` tool);
- [docs/WORD_INTERNALS.md](docs/WORD_INTERNALS.md): read it before changing code or debugging an error.
  It holds everything learned about Word: how it is built, its automation API, the text and paragraph
  model, WordOpenXML, how `Range.InsertXML` resolves styles and lists, the COM error codes, Word's
  surprising behaviours, the server's architecture and invariants, and a symptom -> cause -> fix playbook.

Rules:

- **Never test on a user's real document.** Reproduce with a disposable document
  (`tests_fixtures.build_integration_doc`, or `Documents.Add(Visible=False)` closed with `Close(0)`).
- Paragraph numbers come only from `paragraph_map()`; every edit target is verified against Word
  (`locate()` / `verify_paragraph()`). Keep the invariants listed in WORD_INTERNALS.md section 11.
- Destructive edits refuse to remove pictures, comments, tables or section breaks unless
  `allow_content_loss=true`.
- Never save a document unless the user asked; saving is the user's decision.
- After a change, run every test suite **one at a time** (they share one Word); all must end with
  `ALL CHECKS PASSED`:
  `test_anchors.py`, `test_tables.py`, `test_word_mcp.py`, `test_context.py`, `test_v2.py`, `test_styles.py`,
  `test_concurrent_user.py` (`python -I <suite>`).
- Code must stay valid on Python 3.10 and the MCP SDK 1.x (`mcp>=1.10,<2`).
- When renaming a tool parameter, keep accepting the old name: clients cache tool schemas.
