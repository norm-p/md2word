# md2word -- End-to-End Process Workflow

This document describes the complete md2word pipeline from CLI invocation through
final output, covering **create mode**, **update mode** (AI-augmented), and
**deterministic-only mode** (no API key required).

---

## High-Level Pipeline

![High-Level Pipeline](images/High-Level%20Pipeline-2026-03-26-133243.png)

| Step | Logic | Description |
|------|-------|-------------|
| Interpret | Deterministic | `has_base` = True --> update, False --> create |
| Select approach | Deterministic | `.xlsx` --> excel; update --> xml_edit; else --> pandoc |
| Dispatch | -- | Run selected approach |
| Validate | Deterministic | Open output with python-docx to confirm it is not corrupt |

---

## Create Mode (pandoc.py)

![Create Mode](images/Create%20Mode-2026-03-26-133236.png)

Straight conversion via the bundled `pandoc.exe` (WSL) or system pandoc.
When `--ref-doc` is provided, the output inherits fonts, colors, and named
styles from the reference template. No AI is involved.

**WSL handling**: When pandoc is a Windows `.exe` running under WSL, paths are
converted with `wslpath -w` and pandoc is called directly via `subprocess`
(bypassing pypandoc, which validates `os.path.exists` on Linux paths).

---

## Update Mode (xml_edit.py)

The update pipeline is a multi-step process that preserves the base DOCX's
formatting while applying content changes from the Markdown source.

![Update Mode](images/Update%20Mode-2026-03-26-133210.png)

**Color key**: Blue = inputs, Orange = AI-powered steps, all others = deterministic.

---

## Update Mode -- Detailed Step Breakdown

### Step 1: Conflict Detection (ai/conflict.py) -- Deterministic

Scans `word/document.xml` for `w:ins` (insertions) and `w:del` (deletions)
indicating tracked changes. Counts comments from `word/comments.xml`. If any
are found, they are automatically accepted before proceeding.

### Step 2: Chunking (ai/chunk.py) -- Deterministic

1. **Build heading style map** from `styles.xml` -- identifies paragraph styles
   that function as structural headings (built-in `Heading1-9`, custom styles
   with `w:outlineLvl`, styles inheriting from headings).
2. **Build list style map** -- finds `ListBullet` / `ListNumber` style IDs.
3. **Walk all `w:body` children** -- not just paragraphs. Tables (`w:tbl`),
   structured document tags (`w:sdt`), and other elements are included in
   their enclosing section. `w:sectPr` is excluded.
4. **Group into `DocxSection` objects** -- each section has a heading, level,
   and raw OOXML fragment.

### Step 3: Mapping (ai/map.py) -- AI-Powered

1. **Parse MD sections** via `_parse_md_sections()` -- detects ATX headings
   (`# Title`). Falls back to numbered-list headings if fewer than 3 ATX
   headings found. Bold all-caps lines become boilerplate separators.
2. **LLM alignment** -- the LLM receives MD headings with content previews and
   the DOCX heading list, then produces a mapping: each MD section is marked
   `replace`, `insert`, or `unchanged` relative to a DOCX section.
3. **Post-mapping re-split** -- when a large MD section maps to a small DOCX
   section, scans for unmatched DOCX headings appearing as lines within the
   MD content and splits at those points (no LLM needed).

**Batching**: Documents with >30 MD sections are split into batches of 30 for
the LLM call.

### Step 4: Unchanged Detection (xml_edit.py) -- Deterministic

Two-pass filter applied after mapping:

| Pass | Condition | Result |
|------|-----------|--------|
| 1 | Boilerplate separator heading (dynamic) | Force unchanged |
| 2 | Normalized text similarity >= 0.90 | Mark unchanged |

**Guards** that override unchanged back to replace:
- Heading paragraph has heading style but empty text
- MD has 5+ more bullet items than DOCX

Text normalization: Unicode NFKD, strip markdown formatting, collapse
whitespace, lowercase.

### Step 5: Edit Plan (ai/edit_plan.py) -- AI-Powered

![Edit Plan (AI-Enabled)](images/Edit%20Plan%20(AI-Enabled)-2026-03-26-133227.png)

- Sections >8000 chars are isolated in their own batch with a 32K token budget.
- The LLM receives a **structural summary** of existing XML (paragraph count,
  styles used, table presence) rather than raw OOXML, reducing token usage.
- Each fragment is validated with `lxml.etree.fromstring()` immediately.
- Heading and list style maps are provided so the LLM uses correct style IDs.

### Step 6: Media/Relationship Validation -- Deterministic

Loads `word/_rels/document.xml.rels` and checks that every `r:id`, `r:embed`,
and `r:link` attribute in the edited XML references a declared relationship.
Warns on broken references (non-fatal).

### Step 7: Change Summary (ai/summarize.py) -- AI-Powered

Unless `--accept-changes` is set, the LLM produces a human-readable bullet
list summarizing what changed. Displayed at the CLI for y/n confirmation.

### Step 8: XML Surgery -- Deterministic

The core assembly step, processing edits in **reverse order** to keep element
indices stable.

![XML Surgery](images/XML%20Surgery-2026-03-26-133202.png)

**Element preservation**: The LLM cannot reliably regenerate complex OOXML
structures (images, tables with merged cells/borders). When the LLM produces
no replacement, original elements are rescued and re-appended.

**Deduplication**: Three passes catch LLM-generated content that duplicates
existing paragraphs (within the same section, across preserved sections, or
matching preserved section headings).

**Unchanged section patching**: Sections marked unchanged still receive
targeted fixes -- new bullets inserted, existing bullets updated, minor text
corrections applied, and table rows spliced -- without LLM involvement.

### Step 9: Output -- Deterministic

1. **Repackage DOCX** -- copies all zip entries from the base, replacing only
   `word/document.xml`. Preserves compression methods, timestamps, and entry
   ordering.
2. **Auto-versioning** -- `report.docx` becomes `report_v001.docx`; existing
   `_v001` increments to `_v002`. Existing files are never overwritten.

---

## Create-Mode Fallback

When update mode detects that neither the MD nor the DOCX have section
structure, it automatically falls back to create mode.

![Create-Mode Fallback](images/Create-Mode%20Fallback-2026-03-26-133219.png)

**Trigger conditions** (all must be true):
- DOCX has exactly one section: `(preamble)` (no heading styles)
- MD has no ATX headings (mapping produces at most 1 entry)
- MD text is substantially larger than DOCX text (>1.5x or >2000 chars more)

---

## Deterministic-Only Mode (No API Key)

When no AI provider is configured (no `AI_MODEL` / API key in `.env`, or the AI
SDK is not installed), update mode runs without any LLM calls.

![Deterministic-Only Mode](images/Deterministic-Only%20Mode-2026-03-26-133152.png)

**What works without AI:**
- Heading-based section matching (exact, case-insensitive, and fuzzy)
- Text similarity detection (preserves sections with >= 90% match)
- Bullet insertion and text corrections within matched sections
- Table row updates, insertions, and stale row removal
- All formatting, images, and styles are preserved

**What requires AI:**
- Semantic section alignment (e.g., "Intro" matching "Introduction")
- OOXML fragment generation for sections with major content changes
- Human-readable change summaries

---

## AI vs Deterministic Steps

| Step | Type | Module |
|------|------|--------|
| 1. Conflict detection | Deterministic | ai/conflict.py |
| 2. Chunking | Deterministic | ai/chunk.py |
| 3. Section mapping | **AI** | ai/map.py |
| 4. Unchanged detection | Deterministic | xml_edit.py |
| 5. Edit plan generation | **AI** | ai/edit_plan.py |
| 6. Media validation | Deterministic | xml_edit.py |
| 7. Change summary | **AI** | ai/summarize.py |
| 8. XML surgery | Deterministic | xml_edit.py |
| 9. Repackage + validate | Deterministic | xml_edit.py |

**Design principle**: AI is used only where semantic judgment is needed (section
alignment, OOXML generation, change summarization). Everything else --
mode detection, conflict scanning, XML manipulation, validation -- is
deterministic Python with lxml.

---

## Error Handling and Recovery

![Error Handling and Recovery](images/Error%20Handling%20and%20Recovery-2026-03-26-133142.png)

The pipeline is designed to degrade gracefully. Individual section failures
produce warnings but do not block the rest of the document. Only whole-document
XML corruption is fatal.

---

## Key Data Structures

### DocxSection

Produced by chunking (step 2). Represents one heading-delimited section of the
base DOCX.

| Field | Type | Description |
|-------|------|-------------|
| `heading` | str | Text of heading paragraph (`(preamble)` for content before first heading) |
| `heading_level` | int | 1--9 for headings, 0 for preamble |
| `xml_fragment` | str | Raw OOXML for all elements in the section including the heading |

### SectionMapping

Produced by mapping (step 3). Links one MD section to its DOCX counterpart.

| Field | Type | Description |
|-------|------|-------------|
| `md_heading` | str | Heading text from the Markdown source |
| `md_heading_level` | int | 1--9 for headings, 0 for preamble |
| `md_content` | str | Full Markdown body text of the section |
| `docx_section` | DocxSection or None | Matched DOCX section (None for inserts) |
| `action` | str | `replace`, `insert`, or `unchanged` |

### Edit

Produced by edit plan (step 5). One content operation to apply.

| Field | Type | Description |
|-------|------|-------------|
| `kind` | str | `replace`, `insert`, or `delete` |
| `target_heading` | str | DOCX section heading to modify (or `(end)` for append) |
| `content` | str | Valid OOXML paragraphs (`w:p` elements) |

---

## Glossary

| Term | Definition |
|------|------------|
| **ATX heading** | A Markdown heading using `#` syntax (e.g., `## Section Title`). Named after the original ATX specification. |
| **Base DOCX** | The existing Word document provided as the second CLI argument. Its formatting and structure are preserved during update mode. |
| **Boilerplate separator** | A bold, all-caps paragraph (8--80 characters) that acts as a structural boundary in the DOCX but is not a true heading style. Examples: `ASSUMPTIONS`, `SCOPE OF WORK`. |
| **Chunking** | The deterministic process of splitting `document.xml` into heading-delimited sections by walking all `w:body` children. |
| **Create mode** | Pipeline mode when no base DOCX is provided. Converts Markdown to a new DOCX via pandoc. |
| **Dedup pass** | Post-edit sweep that removes paragraphs duplicated by the LLM -- within the same section, across preserved sections, or matching preserved headings. |
| **DocxSection** | Data structure representing one heading-delimited section of the base DOCX, including heading text, level, and raw OOXML fragment. |
| **Edit** | A single content operation (replace, insert, or delete) targeting a named DOCX section, containing valid OOXML content. |
| **Edit plan** | Step 5 of update mode. The LLM generates OOXML replacement fragments for each section that needs changes, batched and validated. |
| **Element rescue** | When the LLM produces no images or tables for a section that originally had them, the original elements are re-appended after the new content. |
| **Heading style map** | A dictionary mapping DOCX paragraph style IDs to heading levels (1--9). Built from `styles.xml` using pattern matching and inheritance resolution. |
| **List style map** | A dictionary identifying the `ListBullet` and `ListNumber` style IDs used in the document. |
| **lxml** | Python XML library used for all namespace-safe XML parsing and serialization throughout the pipeline. |
| **Mapping** | Step 3 of update mode. The LLM aligns each MD section to a DOCX section as `replace`, `insert`, or `unchanged`. |
| **Normalized text** | Text processed through Unicode NFKD, markdown stripping, whitespace collapsing, and lowercasing. Used for similarity comparison. |
| **OOXML** | Office Open XML -- the XML-based format inside `.docx` files. Key namespace prefix: `w:` (word processing). |
| **Preamble** | Content before the first heading in either the MD or the DOCX. Represented as a synthetic section with heading `(preamble)` at level 0. |
| **pypandoc** | Python wrapper for pandoc. Used in create mode on non-WSL systems. |
| **Reference doc** | A DOCX template passed to pandoc via `--reference-doc` that provides fonts, colors, and named styles for the output. |
| **Repackage** | The final step that creates the output DOCX by copying all zip entries from the base and replacing only `word/document.xml`. |
| **Re-split** | Post-mapping optimization: when a large MD section maps to a small DOCX section, splits the MD at lines matching unmatched DOCX headings. |
| **SectionMapping** | Data structure linking one MD section to its DOCX counterpart with an action (`replace`, `insert`, `unchanged`). |
| **Similarity threshold** | 0.90 -- sections with >= 90% normalized text similarity are marked unchanged and skip LLM processing. |
| **Update mode** | Pipeline mode when a base DOCX is provided. Preserves the base document's formatting while applying content changes from the Markdown source. |
| **w:body** | The root element of `document.xml` containing all document content (paragraphs, tables, structured document tags). |
| **w:drawing** | OOXML element containing an embedded image. Rescued during element preservation when the LLM does not reproduce it. |
| **w:p** | OOXML paragraph element. The most common child of `w:body`. |
| **w:sectPr** | Document section properties element (page size, margins, headers/footers). Excluded from chunking and preserved during repackaging. |
| **w:tbl** | OOXML table element. Complex structure with merged cells, borders, and column widths that the LLM cannot reliably regenerate. |
| **w:tr** | OOXML table row element. Targeted row insertion patches individual rows rather than regenerating entire tables. |
| **XML surgery** | Step 8 of update mode. Edits are applied to `document.xml` in reverse order, followed by deduplication, style injection, and validation. |
| **Zip surgery** | Repackaging the output DOCX by writing modified XML directly into the zip archive, preserving all other entries, compression methods, and timestamps. |
