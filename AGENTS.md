# Agents working on word-mcp

Read [docs/WORD_INTERNALS.md](docs/WORD_INTERNALS.md) before changing code or debugging an error. It holds
everything reverse engineered about Word for this project: how Word is built, its automation API, the
text and paragraph model, WordOpenXML, how `Range.InsertXML` resolves styles and lists, the COM error
codes, Word's surprising behaviours, the server's architecture and invariants, and a symptom -> cause ->
fix playbook.

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
  `test_anchors.py`, `test_word_mcp.py`, `test_context.py`, `test_v2.py`, `test_styles.py`,
  `test_concurrent_user.py` (`python -I <suite>`).
- Code must stay valid on Python 3.10 and the MCP SDK 1.x (`mcp>=1.10,<2`).
- When renaming a tool parameter, keep accepting the old name: clients cache tool schemas.
