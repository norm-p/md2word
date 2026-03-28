# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Source File Safety

**Never modify the original source MD or base DOCX files.**

- Do NOT use `sed`, `awk`, `echo >`, or any shell command that writes to source files
- Do NOT use Edit/Write tools on source MD files (e.g., `_v2.md`) or base DOCX files
- Shell commands that *read* source files as arguments (e.g., `uv run md2word SOURCE.md BASE.docx`) are fine
- All writes go to output files (`_vNNN.docx`, `_vNNN.md`, review docs, code under `src/`)

---

## Repository Layout

```
md2word/
├── src/md2word/           # main package
│   ├── cli.py             # CLI entry point
│   ├── validate.py        # output validation
│   ├── ai/                # AI-assisted modules (client, chunk, map, edit_plan, etc.)
│   └── approaches/        # conversion approaches (pandoc.py, xml_edit.py)
├── samples/               # sample MD + DOCX pairs for manual testing
├── templates/             # (empty) placeholder for --ref-doc templates
├── tools/                 # vendored binaries (pandoc.exe) + validation scripts
│   └── validate_html.py   # pandoc HTML validation for DOCX output (standalone)
├── images/                # diagrams and logo for README/docs
├── dev-docs/              # internal development notes (iteration process, reviews, workflow)
├── dist/                  # built artifacts (md2word.exe)
├── build_exe.ps1          # PyInstaller build script (Windows)
├── md2word.spec            # PyInstaller spec file
├── freeze_hook.py         # PyInstaller runtime hook
├── md2word_entry.py       # PyInstaller entry point wrapper
├── .github/               # copilot-instructions.md
└── .archive/              # archived files
```

**No automated tests exist yet.** `uv run pytest` is the intended command when tests are added.
Manual testing uses the sample files in `samples/`.

## Commands

```bash
uv sync                              # install/update dependencies (required after adding lxml)
uv run md2word --help                 # show CLI usage
uv run md2word INPUT.md               # create mode (new .docx)
uv run md2word INPUT.md BASE.docx     # update mode, writes BASE_v001.docx
uv run md2word INPUT.md BASE.docx --accept-changes  # skip confirmation
uv run md2word INPUT.md BASE.docx --overwrite        # overwrite base in place (no _vNNN)
uv run md2word INPUT.md BASE.docx -o OUT.docx -v    # explicit output, verbose
uv run pytest                         # run tests (none exist yet; use samples/ for manual testing)
```

## Environment

- Python 3.12 (`.python-version`), managed with **uv**
- Pandoc auto-downloads on first run if not found (via `pypandoc.download_pandoc()`)
- `PYPANDOC_PANDOC` in `.env` overrides auto-detection if you have a specific binary
- Copy `.env.example` to `.env` and set `AI_PROVIDER`, `AI_MODEL`, and the matching API key
- AI is optional -- the tool works without any API key in deterministic-only mode
- `lxml` is a required dependency -- run `uv sync` after first clone or after pulling changes

## Architecture

The app converts Markdown to Word documents. Mode is determined deterministically: base provided = update, otherwise = create.

```
INPUT.md + [BASE.docx]
  1. interpret   (det)  -- create or update based on whether base is provided
  2. select      (det)  -- pick approach: pandoc or xml_edit
  3. dispatch           -- run the chosen approach
  4. validate    (det)  -- open output with python-docx to confirm not corrupt
```

### Create mode (`approaches/pandoc.py`)
Straight pypandoc conversion. Supports `--ref-doc` for template styling, `--toc` for table of contents.

### Update mode (`approaches/xml_edit.py`)
Multi-step pipeline that preserves the base DOCX's formatting:
1. `ai/conflict.py` -- scan for tracked changes/comments; accept if found (deterministic)
2. `ai/chunk.py` -- split `word/document.xml` into sections by heading (deterministic)
3. `ai/map.py` -- AI aligns MD sections to DOCX sections (insert/replace/unchanged)
4. `ai/edit_plan.py` -- AI generates OOXML replacement fragments, batched + validated
5. Relationship ID check -- warn on broken media references (deterministic)
6. `ai/summarize.py` -- AI produces human-readable change summary for y/n confirmation
7. XML surgery: patch `document.xml`, rezip preserving entry metadata

`--accept-changes` bypasses step 6.

### Chunking (`ai/chunk.py`)
Iterates ALL direct `w:body` children — `w:p`, `w:tbl`, `w:sdt`, etc. — not just paragraphs.
Tables and structured-document tags are included in their enclosing section's XML fragment so
the AI receives the full section content. `w:sectPr` (document section properties) is excluded.
Uses `lxml` for namespace-safe parsing and serialization.

### Edit plan batching and validation (`ai/edit_plan.py`)
- Processes sections in batches of 5 to stay within LLM output token limits
- Sends a structural summary of existing XML (styles, paragraph count, table presence) instead of
  raw XML to reduce token usage and improve accuracy
- Validates each fragment immediately with `lxml.etree.fromstring()`
- On invalid XML: retries the single failing section at temperature=0 with the parse error included
- On second failure: emits a warning and preserves the original DOCX content for that section
- `max_tokens=16384` per batch call (configurable via `EDIT_MAX_TOKENS` in edit_plan.py)

### JSON parsing (`ai/client.py`)
`parse_llm_json(raw)` strips markdown fences and handles preamble text before the JSON array.
All AI modules use this instead of raw `json.loads`. On decode failure, each call site sends a
retry with the error message before raising.

### Output versioning
When no `-o` is specified in update mode, output auto-versions from the base filename:
`report.docx` -> `report_v001.docx`, `report_v001.docx` -> `report_v002.docx`. Existing files are skipped.

### LLM client (`ai/client.py`)
All AI modules call `LLMClient.complete(system, messages, max_tokens=8192) -> str`.
The client is provider-aware -- set `AI_PROVIDER` to `anthropic`, `azure`, `bedrock`, `vertex`,
`openai`, or `azure-openai`. Modules never touch the SDK directly.

### XML surgery / element preservation (`approaches/xml_edit.py`)
`_apply_edits()` processes edits in reverse order to keep indices stable. For `replace` edits it
rescues two categories of elements before removing the old section, then re-appends them after
the LLM-generated content if the LLM produced none:

- **Drawings** (`w:drawing` inside `w:p`) — images whose relationship IDs are preserved in the zip
- **Tables** (`w:tbl`) — complex OOXML table structures (merged cells, borders, column widths)
  that the LLM cannot reliably regenerate from Markdown table text

Re-appended elements are placed after new content in this order: drawings first, then tables.
A warning is echoed to the console when tables are rescued.

### Unchanged detection (`approaches/xml_edit.py`)
`_sections_text_match()` compares normalized text between MD and DOCX. Threshold is **0.90**
(sections with ≥90% text similarity are marked unchanged and never sent to the LLM). This
protects large table-heavy sections (e.g., Assumptions at 0.93, Appendix B at 0.94) from LLM
regeneration.

### Zip repackage (`approaches/xml_edit.py`)
Uses `ZipFile.writestr(ZipInfo, data)` to copy entries directly from source to output, replacing
only `word/document.xml`. This preserves original compression methods, timestamps, and entry
ordering rather than re-extracting to a temp directory.

### Key design principle
AI is used only where semantic judgment is needed (section alignment, edit generation, change
summarization). Everything else -- mode detection, conflict scanning, XML manipulation,
validation -- is deterministic Python. `lxml` is used throughout for all XML operations.

### Known issue: large MD section mapped to small DOCX section (Issue 2)
Some DOCX documents use custom heading styles for sub-sections that pandoc does not export as
ATX headings in the source MD. The user then authors the MD with those headings as plain body
text lines. The mapper sees one large MD section and one small DOCX section; sub-sections in
the DOCX go unmatched; the LLM cannot regenerate the full content; severe content loss results.

**Do NOT fix this in `_parse_md_sections()`.** Heuristics on the MD side (all-caps detection,
minimum length, blank-line guards) are document-specific and produce false positives or require
level-assignment guesses with no ground truth.

**Recommended fix — post-mapping re-split in `map.py`:**
After `map_sections()` returns, scan for `replace` mappings where text similarity is very low
(< 0.30) and the MD section is much larger than the matched DOCX section. For each such
mapping, inspect the list of **unmatched** DOCX sections and check whether any of their heading
texts appear as a full line (case-insensitive) within the large MD section's content. Where
they do, split the MD content at those positions into synthetic MD sections and map each one
directly to its DOCX section — no LLM call required, the match is exact. The original mapping
becomes just the intro text above the first split point.

This approach is document-agnostic: detection is driven by what headings the DOCX actually has
(ground truth), not assumptions about MD formatting. Each synthetic section inherits its heading
level from chunk.py's DOCX section data. It is self-limiting — only fires on the specific
mismatch pattern — and requires no changes to `_parse_md_sections()`.

Implement as a post-processing step inside `map.py` (not in `xml_edit.py`). The re-split is
conceptually part of the mapping problem and keeping it in `map.py` maintains the clean
separation between mapping and XML surgery.

### Known issue: table preservation for sections with genuine row changes (Issue 4)
Fix B2 re-appends original `w:tbl` elements when the LLM generates no tables. For sections
where the MD adds new table rows (e.g., Appendix E revision history), the right fix is
targeted `w:tr` insertion: preserve the original table structure and splice in only the new
rows. The LLM can assist in a constrained way — given only the table XML and the new row
content, ask it to produce the additional `w:tr` fragments — rather than regenerating the
entire section.

### Round-trip MD validation limitations
The round-trip `.md` is generated by MarkItDown (mammoth → markdownify). This pipeline has
known lossy behaviors that can make the round-trip MD diverge from the source MD even when
the DOCX is correct:
- **Bold boundaries** reflect OOXML run structure, not visual appearance. `**label**:` vs
  `**label:**` depends on which `w:r` the colon lands in — both may look identical in Word.
- **Whitespace collapsing** — HTML intermediate stage collapses double spaces to single.
- **Heading levels** come from style name only (`Heading 1` → `h1`), ignoring list numbering.
  A numbered appendix heading may render correctly in Word but show as `# H1` in round-trip.
- **Smart quotes** pass through as-is from the DOCX; they are not normalized.

When a round-trip comparison flags a discrepancy in bold scope, whitespace, heading level,
or quote style, **visually inspect the DOCX in Word** before treating it as a code defect.

### Edit plan prompt gaps (known)
The LLM prompt in `ai/edit_plan.py` has these known gaps that contribute to open issues:
- No bold-scope instruction — shows `<w:b/>` syntax but no rule about matching `**` markers
- No duplicate-heading prevention — LLM can emit the section heading twice
- No smart-quote guidance — LLM may normalize or alter quote characters
- Table reconstruction is blind — LLM gets `has_tables: true` + 150-char preview, not structure
