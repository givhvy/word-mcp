---
name: m8-word-docs
description: How to edit M8 Word documents (review comments, integration overview, design docs) with the word-new MCP so the content is correct, minimal and verified. Use for any Word task in the Mission 8 workspace.
---

# M8 Word document work

## Rules
- Use only the `word-new` MCP. Never the old `word` MCP.
- Edit the open document in place. Do not rename it (SharePoint). Do not save, resolve comments or reply to comments unless the user says so.
- Keep the user's wording when they ask for it. Change only what the comment asks for.
- Keep text short and non-technical. Do not apply the unslop skill to document text.

## Workflow for a review comment
1. Read the comment with `word_list_comments`. Note the text, the anchor text (`on`) and the paragraph.
2. If the user sends a screenshot, find what it points at: which table, column or paragraph the comment is attached to, and what the reviewer wants changed. Say it back in one line.
3. Gather facts from code in `D:\M8\Mission8` before writing. Use the `explore` subagent for "which component uses X". Use adapter code for formats and protocols. Use the release 1 design document (DD130) for scope. Never invent links, names or deliverables.
4. Make the smallest edit that solves the comment. Prefer `word_table_edit`, `word_set_paragraph`, `word_replace_section`. Fall back to `word_com` with a `Range(start,end)` after reading the real text.
5. Verify: re-read the changed cells or paragraphs and the neighbours, check the contents page, captions and figure numbers.
6. Report in short bullets: what changed, what was not touched, what is unconfirmed.

## Safety checks
- Before any edit by number, confirm the paragraph text with `word_com` (`Paragraphs(n).Range.Text`) or `word_read`. Use `expect`.
- A wrong target once overwrote a contents line (TOC 4.10). Always read the target and the lines around it first. Undo with `word_undo` if wrong.
- If a call fails with "comment card is selected", run `Documents("<url>").Range(0,0).Select` through `word_com`, then retry.
- If `word_list_comments` or `word_save` returns `<unknown>.Name`, retry.
- Comment numbers shift after resolving. List them again before each step.

## Content rules for the integration overview (D0180)
- Separate implemented, optional and mock-backed integrations.
- Release 1 scope: Pre-Race Recons and Parcour. Radio, Velon and WhatsApp are not in DD130, so confirm before keeping them.
- Overview table column is "Component". Use the Angular component names from `M8.UI.Modules/src/ui/src/modules`.
- Per-integration tables have Purpose, Protocol, Format, Data received, Component. The Component row repeats the overview table cell for that integration, so change both together.
- Add a table row with `Tables(t).Rows.Add`, then set the two cells. It copies the last row's formatting. Check the table by its Purpose text and row count before adding.
- Formats from adapter code: JSON for VeloViewer, Open-Meteo, Windy, Velon, Feniks; RainViewer JSON manifest plus PNG tiles; radio WAV/WebM in and JSON out; Street View is an embed iframe.

## Component map (verified in code)
- VeloViewer: Stage Detail, Critical Sections, Stage Reconnaissance (data comes from master data loaded by `M8.Adapter.VeloViewer`).
- Street View: Street View Embed in Stage Detail, Critical Sections, Hazard Detection stage map.
- Open-Meteo, Windy: Weather source toggle, Stage Briefing, Mission Prediction, Critical Sections.
- RainViewer: Stage Map, Rider Position Map (via the `rain_radar_snapshot_feed` feed).
- Feniks AI: Stage Narrative briefing deck, Briefing Modal, Stage Briefing AI Insights (via Aurora `aur_stage_narrative`).

## Techniques that worked
- A comment on one heading can apply to every section ("for each integration"). Re-read the comment, then apply it to all sections, and list which ones could not be done.
- Check a reference exists before linking to it. The References table (table 5) holds the real deliverable names and URLs. Read a link with `Tables(5).Range.Hyperlinks(i).Address`. To check what a deliverable covers, `word_open` it with `read_only:true`, use `word_find` for the integration name, then close it with `Documents("<name>").Close(0)` through `word_com`.
- Add a paragraph after a caption: check the caption prefix, `Paragraphs(n).Range.InsertParagraphAfter`, set style `Body Text`, then `Range(start,start).InsertAfter(text)`. Work bottom-up so paragraph numbers stay valid.
- Make a clickable link: `word_com` cannot pass a Range, so call `Range(a,b).InsertXML` with a flat-OPC package holding `<w:fldSimple w:instr=' HYPERLINK "url" '>`. Copy the URL form (`%20`, `?web=1`) from the References table. Then set `Range(a,b).Style = "Hyperlink"` on the link text. `fldSimple` links come out as plain text, with no blue and no underline. Check `Range.Hyperlinks.Count`, `Font.Color` (16711680 = blue) and `Font.Underline` (1).
- If the user's screenshot is missing or unreadable, say so and list the unresolved comments with `word_list_comments` instead of guessing.
- Do not send guessed arguments to `word_com` (a stray `Hyperlinks.Add` call). Check the result of every call.
- Tables made by `word_replace_section` or `word_insert_markdown` can come out with cell paragraphs in `Cover - Project Info` (7 pt, no space after), unlike the document's own tables (`Normal`, 9 pt, 6 pt after). After inserting, compare `Tables(n).Cell(2,1).Range.Font.Size` with an original table, and fix with `word_com` `Tables(n).Range.Style` set to `Normal`.
- To read a whole document, call `word_read` in chunks of about 150 paragraphs and return the text without `.slice`. `word_context` and long tool results get cut with "[truncated]", which is a harness output limit, not a Word MCP limit. `word_find` paragraph numbers can be 1 higher than `word_read` numbers after edits, so use the `word_find` number with `word_set_paragraph` plus `expect`.
- `word_replace` fails with "String parameter too long" when `find` is over 255 characters (Word Find limit). Use `word_set_paragraph` with `expect` instead. `word_set_cell` does not parse markdown, so `**` shows literally.

## Figure and table numbers
- Captions are `SEQ Figure` and `SEQ Table` fields. After deleting or adding a figure, update every SEQ field: loop `Fields(i)`, read `Code.Text`, call `Fields(i).Update` through `word_com`, then re-read the captions.
- Read captions from Word, not from `word_read`: its paragraph numbers and styles drift after table edits. Loop `Paragraphs(i).Range.Text` and match `^(Figure|Table) \d+`.
- Also search the body for typed references such as "see Figure 3". The D0180 overview has none, and no list of figures.
- Check that figure numbers run 1..N in order and that each figure has one caption.

## Pulling repos before writing a design document (DD120 and similar)
- The user names the repo, and may send an image of it that does not arrive. If a screenshot is missing, ask which repo, do not guess.
- Pull: `git fetch`, then check `git status`. If there is uncommitted work, `git stash push -u`, `git merge origin/main`, resolve, commit the merge, then `git stash pop`. Keep both sides: "my code" and "the team's code" both stay.
- Typical conflict: both sides append lines at the same spot (csproj lists, `.sln` project blocks and config blocks). Keep all lines. In a `.sln`, put `EndProject` between the two project blocks. Then check: no duplicate GUIDs, project count equals EndProject count, no conflict markers left.
- After a merge in `M8.Library.MessagingContracts`, check each `Schema/M8.Dto.*.yaml` has a csproj, a sln entry and a pipeline entry, then run `dotnet build Solution/MessagingInterfaces.sln -c Release`.
- If the repo only has uncommitted edits and is behind with no local commits (for example `M8.Infrastructure.EnvironmentConfig`), use `git stash push`, `git merge --ff-only origin/main`, `git stash pop`. A list that both sides extended (Aurora `AdapterMessageTypes` in `values.yaml`) conflicts too: keep all entries and check there are no duplicates.
- The word-mcp venv and the system Python have no `yaml` module. Check YAML by reading the merged lines and their indentation, or install a parser only if the user agrees.
- `M8.Configuration.PP`: the team may delete files the user edited (stash pop shows `deleted by us` / `DU`). Before merging, compare `git diff --name-status HEAD...origin/main` with `git status --short`. Save the user's version of deleted files as a patch (`git diff 'stash@{0}^1' 'stash@{0}' -- <path>`, quote the stash ref in PowerShell), then accept the deletion with `git rm`. Check that kept edits still build (`dotnet build BusinessRules/BusinessRules.sln -c Release`) and `Schema/init.json` still parses.
- `M8.Configuration.Initialize`: `MasterData/Schema/master_data_schema.json` and `Permissions/Data/permissions_configuration.json` give many interleaved conflicts, so a line merge breaks the JSON. Merge by structure: take `origin/main`'s file, then add my items by `internalName` (schema `entities`) or `id` (permissions lists). First check with base/ours/theirs that no item was changed on both sides. Write JSON with indent 2, LF, `ensure_ascii=False`.
- Do not pipe `git show` through PowerShell for files with non-ASCII text: the em dash becomes `ΓÇö`. Use `cmd /c "git show ref:path > file"` and read bytes in Python. After writing, diff against `origin/main` and expect 0 removed lines.
- Check after a structural merge: no dangling `entityInternalName`, no duplicate ids, every entity permission points to an existing permission set, all `QueryStream/Feeds/*.xml` parse.
- `M8.Configuration.Aurora`: after the merge, a clean git merge is not enough. Every test `TestDbContext : IAURDbContext` must list all `DbSet`s and entity configs. When both sides add an entity, build fails with `CS0535 does not implement IAURDbContext.<X>` in many test files. Add the missing `DbSet` and `modelBuilder.Entity<...>` block to each. Run `dotnet test Solution/Configuration.Aurora.sln` before and after restoring the stash.
- Aurora `aodb.schema.xml` is CRLF in the working copy and LF in HEAD. A stash pop then conflicts on the whole file. Fix: `git checkout HEAD -- <file>`, build a patch of only the user's additions (convert stash base and stash to LF, `git diff --no-index`), then insert each hunk by its context lines and parse the XML. Write back with the file's original EOL.
- In `cmd /c "git show stash@{0}^^1:path"` the caret must be doubled. In PowerShell write `'stash@{0}^1'`.
- Directory.Packages.props: take the team's DTO package versions, keep only the packages the user added.
- Adapter repos now use central package versions (`Solution/Directory.Packages.props`). In `.csproj`, `PackageReference` has no `Version=`. When the user's local-DTO `Condition` conflicts with that, keep the team's versionless form and add only the user's `Condition`.
- `git stash pop` can leave a file with several `<<<<<<< Updated upstream` blocks (an unrelated `using` line plus a code block). Count the blocks first and resolve each by index in the script. If an assert fails after `git stash drop` already ran in the same command, the file keeps its markers; re-run the fix, do not chain `stash drop` after a script that can fail.
- Do not push unless asked. After a repo is done, ask the user which repo is next.
- Do not edit generated files.
- A DD document is opened read-only first (`word_open`, `read_only:true`) and closed after reading if it was not already open.

## Keep this skill current
- After every Word task, add any new technique, trap or fact to this file before the final reply. Edit existing lines instead of piling up new ones. Remove lines that turned out wrong.
- Put facts about the documents (what exists, what is missing) under "Content rules" so the next task does not re-check them.
- Tell the user in one bullet what was added.

## Tool hygiene
- Send every tool call with all its required parameters. Empty calls fail and waste turns.
- Read the target paragraph and its neighbours before and after each edit. Do not trust paragraph numbers from an older read.
- Close any document you opened for checking, and do not leave it modified.

## Facts about the D0180 overview (as of the last task)
- References table (table 5) has deliverables for: Auto Transcription, Google Street View, Miscellaneous, OpenAI, VeloViewer, Weather & Aerodynamics, Weather APIs, WhatsApp, Windy.
- No deliverable exists for Velon, Open-Meteo, RainViewer or Feniks AI. The OpenAI deliverable covers audio transcription and entity extraction only.
- Release 1 scope is VeloViewer, Street View, Open-Meteo, Windy, RainViewer, Feniks AI. The Velon and Radio sections, their overview rows and the sentences about them were removed (DD130 does not mention them).
- Closing "For further details..." sentences with links are in the VeloViewer, Street View and Windy sections. Open-Meteo, RainViewer and Feniks AI have no deliverable, and the user has not chosen the wording.
- Screenshots were removed. Figures now run 1 to 7 (1 overview, 2 to 7 flow diagrams). Contents page: 4.1 Overview to 4.8 Conclusion.
- Street View uses both the Embed iframe and the Static API still image (briefing deck slide `phase-detail-slide.component.ts`).
- Velon adapter exists (`M8.Adapter.RiderPosition`) but `GraphQLHost` and `ApiKey` are empty by default.
- Do not state a file format (such as JPEG) unless it is in the code.

## Diagrams and pictures
- Figure 1 (System context) is generated by `skills/m8-word-docs/make_context_diagram.py <out.png>` (Pillow, 1600x800, house colours #e4f3f3 fill, #356c78 lines, #18343e text). Edit the `sources` list there when the scope changes, run it, and view the PNG before inserting.
- Do not use `word_insert_image(replace=<n>)`. It crashed Word once (wwlib.dll access violation) and can delete the paragraph before the picture. Insert the new picture after the old one, read it back with `word_list_images`, then delete the old picture's paragraph with `word_delete_paragraphs(allow_content_loss=true)`.
- Check crop, size and `Information(3)` (page) of heading, picture and caption afterwards. A picture that replaced an old one can lose its crop.
- Extract an existing picture for reference with `Paragraphs(n).Range.WordOpenXML` and base64-decode the `pkg:binaryData` of `/word/media/*`.
- When scope changes, re-check the picture. Pictures hold text that `word_find` cannot see.

## Deleting a section
- Deleting a range deletes the comments anchored in it. List the comments in the range first (`Range(a,b).Comments.Count`) and tell the user which ones go.
- Delete bottom-up. Check the first characters of the start and end paragraphs, then delete `Range(startOfHeading, startOfNextHeading)`.
- Afterwards remove its row from any overview table, remove sentences about it elsewhere (introduction, patterns, conclusion), update SEQ fields and `TablesOfContents(1).Update`, then re-read the headings, captions and contents.

## When the MCP fails
Fix it in `D:\wordfolder\word-mcp`: read `AGENTS.md` and `docs/WORD_INTERNALS.md`, reproduce on a disposable document, add a test, run the suites with `.venv\Scripts\python.exe -X utf8 -I <suite>`, commit locally and do not push unless asked.

## Large generated content (DD120 lesson)
- Generate big tables from the schema files with a script, not by hand. DD120 uses `master_data_schema.json`, `init.json` and `aodb.schema.xml`. Dump them to one `model.json`, then build the markdown from it.
- `execute` has `fetch`. Serve the generated files with `python -m http.server <port> --bind 127.0.0.1`, fetch the text inside `execute`, and pass it to `word_insert_markdown`. Stop the server afterwards.
- On a document with a TOC, the tool's paragraph numbers can run 14 or more below Word's. `word_set_paragraph` with an anchor then overwrites TOC lines. Write with `word_com` instead: check `Paragraphs(i).Range.Text` starts with the text you expect, then set `Range(start, end-1).Text`.
- `word_find` and anchors match the TOC entry first. A heading that also sits in the TOC cannot be used as an anchor. Rename it with `Range(s, e-1).Text` set.
- `word_insert_markdown` with `position:"end"` is the safe way to add a large block. Pass `style_overrides:{table:"Table Grid", table_text:"Body Text"}` when the document's own table style is a cover style.
- To place text at a spot, select it with `word_com` (`Paragraphs(n).Range.Select`, then `Selection.Collapse` with 1), then use `position:"cursor"`. `word_insert_image` accepts `cursor` too.
- Undo reverts one tool call. Always read back the paragraph after `word_set_paragraph`.
- Pulse storage facts (verified): Master Data = `<entity>` plus `<entity>_registration`, Number is numeric, DateTime is timestamptz, Lookup is uuid, FileAttachment is two text columns `__file_name` and `__file_hash`. Aurora = one partitioned `public."Entities"` (PK Type+Id, jsonb `AttributeValues`), partition `E_AUR_<Entity>`, DateTime casts are disabled in indexes. PP = `EntityValues` plus `EntityAttributeValues` with one column per type and `Json` for complex types. `pulse_all` is old; Aurora is now the main store.
- E/R diagrams: `npm i @viz-js/viz @resvg/resvg-js` in a temp folder, build DOT with HTML-table nodes, render to PNG. Split large models into groups so the text stays readable at 16 cm.
- When two open documents share a name (local copy and SharePoint copy), a short name resolves to the SharePoint one. Always pass the full local path as `document`.
- `word_read_tables` with `near` fails when the text matches several tables. Pass `table=N`. Table numbers shift after you delete tables.
- `word_delete_paragraphs` on a range that holds tables needs `in_tables:true` and `allow_content_loss:true`. It returns `removed_content`. Check the start and end paragraph text first and delete bottom-up.
- `word_replace` fails when `find` is over 255 characters. Use placeholder-chained replacements.
- E/R diagrams without Node: `D:\DD120 - Physcial Datamodel\diagrams\make_diagrams.py` (Pillow) draws boxes in a grid and links only in the gaps between neighbouring boxes. It stops with an error if a link would cross a box, because the picture cannot be viewed from the tool. Swap pictures with insert-after-then-delete (see "Diagrams and pictures"). The user's DD120 rules: no title or legend text inside the picture, no UQ (PK and FK only), blue entity boxes only, and a short text under each heading that explains colours, PK, FK and arrows.
- Never use `replace=n` on pictures. After swapping, re-read the paragraphs around the picture.
- The Aurora partition name is `E_{Namespace}_{Name}` in `OperationalSchemaUpdater.cs`. The M8 schema namespace is `AUR`, so `E_AUR_<Entity>` is right. A subagent flagged it as wrong because it missed the namespace.

## Replacing a whole catalogue (D0130 lesson)
- `word_replace_section(anchor, markdown, keep_heading=true)` swaps a heading and everything under it in one call, 1,569 paragraphs and 42 tables in about 4 seconds. The result reports removed pictures, comments and tables. Read it.
- It picks the document's cover table style ("Blank" plus "Cover - Project Info", 7 pt) for tables. Pass `style_overrides:{table:"Table Grid", table_text:"Body Text"}`. Undo reverts the whole call, so a wrong style costs one undo and one retry.
- After a replace on a document with a TOC, run `word_update_fields` and check the TOC numbers before you write "see 2.2.3" in text.
- Where the schema leaves a description empty, DD120 and D0130 repeat the attribute name as a short sentence. Say so in the column conventions. Do not invent meaning.
- D0130 needs UML class diagrams, not E/R boxes. The Netcompany guideline (D0130 template) says to use UML class diagrams with attributes, cardinality, association type (aggregation, inheritance) and value objects marked. Boxes with a restyled E/R generator do not meet that. Do not reuse `make_diagrams.py` for D0130.
- D0130 diagrams: `build\make_uml.js` in `D:\D0130 - Logical Datamodel\build` reads `model.json` (written by `make_catalogue.py`) and draws UML with Graphviz (`@viz-js/viz`) and `@resvg/resvg-js`. Install both in `%TEMP%\opencode\viz`, copy the script there and run `node make_uml.js <model.json> <out folder>`. Stereotypes: `entity`, `value object`, `external`. Open arrow = association, filled diamond = composition, multiplicity at both ends.
- The user wants strict UML class diagrams, so D0130 now uses Mermaid (`classDiagram`), not Graphviz. The Graphviz `make_uml.js` is kept only as the old version. Pipeline in `D:\D0130 - Logical Datamodel\build`: `node make_mermaid.js model.json mm` (one `.mmd` per group of linked classes), `render_mermaid.ps1` (mmdc with Chrome, needs `@mermaid-js/mermaid-cli` and `mmd_config.json` and `puppeteer.json` in `%TEMP%\opencode\mmd`, install with `PUPPETEER_SKIP_DOWNLOAD=true` and point `executablePath` at `C:\Program Files\Google\Chrome\Application\chrome.exe`), `pack_mermaid.py` (packs the groups into pages, prints which classes are on each), `check_mermaid.ps1` (counts classes and relations in the SVGs against the `.mmd`).
- Mermaid limits: stereotype names cannot contain spaces in source, but `<<value object>>` does render as «value object». It has no grid packing, so unlinked classes stack in one column. Split by connected group, then pack the pictures yourself. Set the text size by fixing the packed width (1700 px gives about 7 to 8 pt at 16 cm).
- The D0130 figure captions ("Figure 1: ...") are plain Caption-style text, not SEQ fields. When you add or remove a picture, renumber them by hand with `word_set_paragraph` and read them back.
- `word_insert_image` with `caption` failed with 0x80020009 and still left the picture behind. Add the picture without `caption`, then add the caption with `word_insert` and `style: "Caption"`. Undo and check for a leftover empty paragraph after any failed insert.
- `word_com` reads the ACTIVE document. With DD120 and D0130 both open it returned DD120 text for D0130 paragraph numbers. Call `word_activate` with the full path first.
- D0130 now has ONE Mermaid picture per store (Master Data, PP, Aurora), 16.5 cm wide, no part numbers. Associations are plain lines with no arrowhead, because Mermaid cannot draw UML's open V arrowhead and its `-->` draws a filled one. Attributes have no `+` visibility mark. `pack_mermaid.py` searches for the narrowest width at which the trimmed picture still fits 16.5 x 24 cm, and `DIR=LR` (env var) makes PP and Aurora text larger. The text is small for PP and Aurora (about 5 to 6 pt, my estimate).
- PowerShell 5 `Set-Content -Encoding utf8` writes a BOM, which breaks `exec` and Python source. Write files with `[IO.File]::WriteAllText(path, text, (New-Object Text.UTF8Encoding $false))`.
- A box that mixes association attributes with plain attributes repeats the same fact. Draw associations as lines and keep only plain attributes in the box.
- Check a generated diagram against the schema by counting nodes and edges in the SVG (`<g class="node">` and `<g class="edge">`) and by viewing the PNG.
- Opening a second document in Word makes it the active one. Keep passing the full local path as `document`.
- The `shell`, `edit`, `write` and `read` tools sometimes fail with "malformed JSON" when you send empty parameters. Send every required parameter and retry one call at a time. PowerShell breaks `python -c` one-liners with nested quotes. Write a `.py` file instead.
