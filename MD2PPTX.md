# MD2PPTX: Extending md2word to Support PowerPoint

Research and implementation plan for adding PPTX update mode to the md2word tool.

**Use case:** `Yeti - AI Roadmap Executive Readout_v2.md` + `_v1.pptx` --> `_v002.pptx`

---

## 1. Why PPTX Is Structurally Different from DOCX

### DOCX: One Monolithic Body
```
word/document.xml
  w:body > w:p (paragraph) > w:r (run) > w:t (text)
```
All content lives in a single XML file. Sections are delimited by heading styles.

### PPTX: Per-Slide Files + Shape Layer
```
ppt/presentation.xml         (manifest, references slides by rId)
ppt/slides/slide1.xml        (one file per slide)
ppt/slides/slide2.xml
ppt/notesSlides/notesSlide1.xml  (one per slide with notes)
ppt/slideMasters/...         (master formatting)
ppt/slideLayouts/...         (layout templates: Title Slide, Title and Content, etc.)
```

Each slide's content model:
```xml
<p:sld>
  <p:cSld>
    <p:spTree>            <!-- shape tree -->
      <p:sp>              <!-- shape (title, body, image, etc.) -->
        <p:nvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>   <!-- placeholder type -->
        </p:nvSpPr>
        <p:txBody>        <!-- text body inside shape -->
          <a:p>           <!-- paragraph (DrawingML, not WordprocessingML) -->
            <a:r>         <!-- run -->
              <a:rPr b="1"/>   <!-- bold = attribute, not child element -->
              <a:t>Text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
```

### Key Differences Summary

| Aspect | DOCX | PPTX |
|--------|------|------|
| Content files | 1 (`word/document.xml`) | N (`ppt/slides/slideN.xml`) |
| Content unit | Paragraph (`w:p`) | Shape (`p:sp`) containing text body |
| Text namespace | `w:` (WordprocessingML) | `a:` (DrawingML) |
| Bold | `<w:b/>` (element) | `b="1"` (attribute) |
| Font size | `<w:sz w:val="24"/>` (half-pts) | `sz="1800"` (hundredths of pt) |
| Tables | `w:tbl` direct in body | `a:tbl` inside `p:graphicFrame` shape |
| Images | `w:drawing` in paragraphs | `p:pic` shapes in shape tree |
| Image rels | `word/_rels/document.xml.rels` | `ppt/slides/_rels/slideN.xml.rels` (per-slide) |
| Sections | Heading-style delimited | Slide boundaries (files) |
| Notes | N/A | `ppt/notesSlides/notesSlideN.xml` |
| Tracked changes | `w:ins`, `w:del` | Not supported in PPTX |
| Layouts | Paragraph styles | Slide layouts (Title Slide, Section Header, etc.) |

---

## 2. What the Source MD Looks Like

The Yeti MD uses these conventions (from the PPTX-exported markdown):

```markdown
<!-- Slide number: 1 -->
![YETI Logo](Picture4.jpg)
# AI Roadmap Executive Readout
Yeti Warranty AI Use Case

<!-- Slide number: 2 -->
# Meet the Team
...

### Notes:
Speaker: Norm (Solutions Architect)
This is the crawl architecture...
```

**Slide boundaries:** `<!-- Slide number: N -->` HTML comments
**Slide titles:** `# Heading` (ATX H1) on each slide
**Speaker notes:** `### Notes:` subsection at end of slide content
**Images:** `![alt](filename.jpg)` referencing loose image files
**Tables:** Standard markdown pipe tables
**Footer boilerplate:** `(c)2026 3Cloud, Proprietary & Confidential` + page number on every slide

---

## 3. Current Codebase: What's Reusable vs Word-Specific

~70% of modules are reusable by count, but the 30% that needs rewriting includes the largest
file (`xml_edit.py` at ~2800 lines). The saving grace: PPTX's per-slide file structure
eliminates the hardest part of the DOCX pipeline (heading-based chunking and index-stable
reverse-order patching of a monolithic XML body).

### Fully Reusable (no changes needed)

| Module | Why |
|--------|-----|
| `ai/client.py` | Provider-agnostic LLM wrapper (`complete()` + `parse_llm_json()`). Zero format-specific logic. |
| `ai/summarize.py` | Generates human-readable change summary from mapping + edits. Format-agnostic. |
| `ai/interpret.py` | Mode detection (create vs update). Operates on file extension only. |
| `output.py` | Console output formatting. Generic. |
| Output versioning | `_v001`, `_v002` naming in `cli.py`. Extension-agnostic -- already works for any suffix. |

### Reusable with Modifications

| Module | What stays | What changes |
|--------|-----------|-------------|
| `cli.py` | Mode detection, output versioning, dispatch framework | Add `.pptx` case to `_detect_approach()`. Route to `pptx_edit` approach. |
| `ai/map.py` | Core alignment logic: action types (`insert`/`replace`/`unchanged`/`delete`), batching (30-section limit), deterministic fallback (exact + fuzzy title match), LLM alignment prompt framework | MD parser swapped: `_parse_md_sections()` (heading-based) replaced by `_parse_md_slides()` (split on `<!-- Slide number: N -->`). LLM prompt adjusted to describe slides instead of heading-delimited sections. Deterministic fallback simplified to ordered index + title similarity. |
| `ai/edit_plan.py` | Batching framework (5 sections/call), immediate lxml validation of each fragment, retry-with-error-feedback loop, "preserve on second failure" fallback, structural summary approach (send metadata not raw XML) | Prompt examples rewritten: DrawingML (`a:p`, `a:r`, `a:t`, `a:rPr b="1"`) instead of WordprocessingML (`w:p`, `w:r`, `w:t`, `w:b/`). Structural summary adjusted for shapes/placeholders instead of paragraph styles. Namespace in validation changed. |
| `validate.py` | Validation-by-opening pattern | Add: `python-pptx` `Presentation(path)` check alongside existing `python-docx` `Document(path)` check. |
| `approaches/pandoc.py` | pypandoc wrapper, `--ref-doc` support, auto-download | Ensure `.pptx` output extension is passed through. Add preprocessor to convert `<!-- Slide number: N -->` to `---` and `### Notes:` to `::: notes` blocks for pandoc's PPTX writer. |

### Word-Specific (need PPTX equivalents)

| Module | PPTX equivalent needed | Why it can't be reused directly |
|--------|----------------------|---------------------------------|
| `ai/chunk.py` (331 lines) | **Mostly unnecessary.** PPTX is pre-chunked by slide files. A thin `pptx_chunk.py` reads `ppt/slides/slideN.xml` in presentation order, extracts title via `<p:ph type="title"/>`, concatenates body text, and reads notes from `ppt/notesSlides/`. | DOCX chunking parses `word/styles.xml` for heading styles, walks `w:body` children, handles custom outline levels -- none of which exist in PPTX. |
| `ai/conflict.py` (69 lines) | **Not needed.** PPTX has no tracked changes (`w:ins`/`w:del` are WordprocessingML-only). | N/A -- entire module is skipped. |
| `approaches/xml_edit.py` (~2800 lines) | **New `pptx_edit.py` required.** This is the largest piece of new work. | Every deterministic fix (Fix A heading preservation, Fix B style stripping, Fix B2 duplicate heading removal, drawing/table rescue) is built around `w:p`, `w:r`, `w:tbl` elements. PPTX uses `p:sp`, `a:p`, `a:r`, `p:graphicFrame`/`a:tbl`. The shape-tree model (positioned shapes with EMU coordinates) is fundamentally different from DOCX's flowing paragraph model. |
| Heading style detection | **Replaced by placeholder detection.** Title shapes identified by `<p:ph type="title"/>` attribute, not style inheritance from `word/styles.xml`. | DOCX heading detection walks style inheritance chains and outline levels. PPTX has no equivalent -- slide structure is defined by layout placeholders. |
| `_apply_edits()` reverse-order patching | **Replaced by per-slide patching.** Each slide is an independent file; no index stability concerns. | DOCX patches a single `w:body` in reverse order to keep element indices stable during insert/delete. PPTX edits are isolated per-slide -- modifying slide 5 cannot shift elements in slide 3. |
| Zip repackage (single-file replace) | **Multi-file repackage.** Must replace N `ppt/slides/slideN.xml` files + N `ppt/notesSlides/` files. For slide insert/delete: also update `ppt/presentation.xml`, `[Content_Types].xml`, and `.rels` files. | DOCX repackage swaps exactly one entry (`word/document.xml`). PPTX's distributed file structure means more entries to manage, but each is smaller and self-contained. |
| Unchanged detection (`_sections_text_match`) | **Reusable concept, new implementation.** Same 0.90 similarity threshold, but compare per-slide text (title + body) instead of heading-delimited section text. | Text extraction differs: DOCX walks `w:p/w:r/w:t`; PPTX walks `p:sp/p:txBody/a:p/a:r/a:t`. The normalization and comparison logic is identical. |
| Table/drawing rescue | **Reusable concept, new element types.** Rescue `p:graphicFrame` (tables) and `p:pic` (images) instead of `w:tbl` and `w:drawing`. | Element names and parent structure differ, but the rescue pattern (detect missing elements in LLM output, re-append originals) is identical. |

---

## 4. Architecture for PPTX Support

### 4.1 Create Mode (MD --> new PPTX)

**Already works.** Pandoc natively supports `--to pptx`:
```bash
pandoc input.md -o output.pptx --reference-doc=template.pptx
```

Pandoc auto-selects slide layouts:
- Title metadata --> "Title Slide"
- H1 above slide-level --> "Section Header"
- H2 + content --> "Title and Content"
- Two-column divs --> "Two Content"

Only change needed: `approaches/pandoc.py` already calls pypandoc; just ensure it handles `.pptx` output extension.

**Caveat:** The Yeti MD uses `<!-- Slide number: N -->` comments, not heading levels, to delimit slides. Pandoc ignores HTML comments when producing PPTX. The MD would need preprocessing to insert `---` (horizontal rules) at slide boundaries, or use `--slide-level=0` with `---` inserted. This is a one-time transform in the pipeline.

### 4.2 Update Mode (MD + base PPTX --> updated PPTX)

This is the hard part. The current DOCX update pipeline is:

```
1. conflict scan          --> NOT NEEDED for PPTX
2. chunk by heading       --> REPLACED by per-slide extraction
3. pre-compare            --> REUSABLE (compare round-trip MD vs source MD)
4. map sections (AI)      --> REUSABLE (map MD slides to PPTX slides)
5. edit plan (AI)         --> NEEDS NEW PROMPTS (DrawingML instead of WordprocessingML)
6. summarize (AI)         --> REUSABLE as-is
7. apply edits (det)      --> NEEDS NEW IMPLEMENTATION (per-slide XML surgery)
8. repackage zip (det)    --> NEEDS MODIFICATION (multi-file replace)
9. validate (det)         --> NEEDS python-pptx check
```

### 4.3 PPTX Chunking: Simpler Than DOCX

DOCX chunking (`ai/chunk.py`, 331 lines) parses heading styles from `word/styles.xml` and walks `w:body` children to split by heading boundaries. This is complex because headings are just specially-styled paragraphs.

PPTX chunking is trivial: **each slide is already a separate file.** The equivalent of `DocxSection` is:

```python
@dataclass
class PptxSlide:
    index: int                  # slide number (1-based)
    title: str                  # text from title placeholder shape
    body_text: str              # all non-title text, concatenated
    notes_text: str             # speaker notes (from notesSlide)
    shapes: list[Shape]         # shape metadata for structural summary
    slide_path: str             # "ppt/slides/slide3.xml"
    notes_path: str | None      # "ppt/notesSlides/notesSlide3.xml"
    xml: str                    # raw slide XML
```

Extraction: iterate `ppt/slides/slideN.xml` files in presentation order, parse each for title placeholder (`<p:ph type="title"/>`) and body content.

### 4.4 MD Slide Parsing

The Yeti MD uses `<!-- Slide number: N -->` boundaries. The parser would:

```python
@dataclass
class MdSlide:
    number: int
    title: str          # first # heading after boundary
    body: str           # everything between title and next boundary or ### Notes:
    notes: str | None   # content after ### Notes: if present
    images: list[str]   # ![alt](filename) references
```

This replaces `_parse_md_sections()` in `map.py`. Much simpler since boundaries are explicit.

### 4.5 Mapping (Slide Alignment)

The mapping problem is simpler for PPTX:
- DOCX: fuzzy heading matching because sections are implicit
- PPTX: slides are numbered and titled; alignment is mostly by order + title match

Reuse `map.py` with a PPTX-specific `_parse_md_slides()` and adjust the LLM prompt to describe slides instead of heading-delimited sections. Deterministic fallback: ordered index matching + title similarity.

Actions remain the same: `insert`, `replace`, `unchanged`, `delete`.

### 4.6 Edit Plan: DrawingML Generation

This is the most significant prompt change. The current `edit_plan.py` teaches the LLM to generate WordprocessingML (`<w:p>`, `<w:r>`, `<w:t>`). For PPTX, it must generate DrawingML-in-PresentationML.

**Approach A: Full XML generation (like DOCX)**
- LLM generates complete slide XML: `<p:cSld><p:spTree>...`
- Pro: maximum flexibility
- Con: PPTX shapes have exact positions (EMUs), layout references, and placeholder indices that the LLM would need to preserve. Much more fragile than DOCX paragraph generation.

**Approach B: Text-only generation + python-pptx assembly (recommended)**
- LLM generates structured JSON describing slide content changes:
  ```json
  {
    "title": "Updated Title",
    "body_paragraphs": [
      {"text": "Bullet one", "level": 0, "bold": false},
      {"text": "Sub-bullet", "level": 1, "bold": false}
    ],
    "notes": "Updated speaker notes"
  }
  ```
- Deterministic code uses python-pptx to apply changes to existing shapes, preserving positions/sizes/layouts
- Pro: avoids LLM generating fragile position/size attributes; preserves slide master formatting
- Con: limited to text changes; layout changes (adding/removing shapes) need special handling

**Approach C: Hybrid**
- Use Approach B for text content changes (most common case)
- Fall back to Approach A for structural changes (adding tables, rearranging shapes)
- This mirrors the DOCX tool's philosophy: deterministic where possible, AI only for semantic judgment

**Recommendation: Start with Approach B**, which handles the 80% case (text/bullet updates, title changes, notes changes). Add Approach A capabilities incrementally for tables and structural changes.

### 4.7 Apply Edits: Per-Slide XML Surgery

DOCX `_apply_edits()` patches a single `word/document.xml`. PPTX equivalent patches individual slide files.

For text-only changes (Approach B):
```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation(base_pptx)
slide = prs.slides[index]

# Update title
for shape in slide.placeholders:
    if shape.placeholder_format.type == PP_PLACEHOLDER.TITLE:
        shape.text = new_title
        # Reapply formatting from original runs

# Update body
for shape in slide.placeholders:
    if shape.placeholder_format.idx == 1:  # body placeholder
        tf = shape.text_frame
        tf.clear()
        for para in new_paragraphs:
            p = tf.add_paragraph()
            p.text = para.text
            p.level = para.level
```

For structural changes (Approach A): direct lxml manipulation of `ppt/slides/slideN.xml`, similar to current DOCX XML surgery but with DrawingML elements.

### 4.8 Zip Repackage

Current DOCX repackage replaces only `word/document.xml`. PPTX repackage must:

1. Replace modified `ppt/slides/slideN.xml` files
2. Replace modified `ppt/notesSlides/notesSlideN.xml` files
3. For inserted slides: add new slide XML files, update `ppt/presentation.xml` (slide list), update `[Content_Types].xml`, add new `.rels` files
4. For deleted slides: remove slide files, update presentation.xml, update Content_Types
5. Preserve all other entries (theme, masters, layouts, media)

Insert/delete is significantly more complex than DOCX (which never adds/removes files from the zip). Using python-pptx's `Presentation.save()` would handle this automatically but would lose the byte-level preservation of the current zip approach.

**Recommendation:** Use python-pptx for insert/delete operations (correctness over byte preservation). Use direct zip manipulation only for in-place slide content updates where no structural changes occur.

### 4.9 Image Handling

The Yeti MD references images like `![alt](Picture4.jpg)`. In update mode:

- **Existing images:** Already in `ppt/media/` with relationships in per-slide `.rels`. Preserve as-is for unchanged slides.
- **New images:** Must be added to `ppt/media/`, with a new relationship entry in the target slide's `.rels`. python-pptx handles this via `slide.shapes.add_picture()`.
- **Image rescue:** Analogous to DOCX drawing rescue. If LLM-generated content drops image references, re-append original `p:pic` shapes.

### 4.10 Table Handling

PPTX tables are more complex than DOCX tables:
- Wrapped in `p:graphicFrame` with explicit position/size
- Each cell has full `a:txBody` structure
- Column widths in EMUs (not relative)

**Recommendation:** Same strategy as DOCX -- rescue original tables when LLM cannot reliably regenerate. For row-level changes, use targeted `a:tr` insertion (same as the Issue 4 approach for DOCX).

---

## 5. Differences in the MD Parsing Problem

### DOCX Pipeline
The MD is written by a human editing a pandoc-exported markdown. Heading levels delimit sections. The challenge is that custom DOCX heading styles may not export as ATX headings, leading to the "large MD section mapped to small DOCX section" problem (Issue 2).

### PPTX Pipeline
The MD has explicit `<!-- Slide number: N -->` boundaries. No ambiguity about where slides start/end. The mapping problem is simpler: match MD slide N to PPTX slide N, with title as confirmation.

However, new challenges arise:
- **Slide content is spatial, not linear.** A DOCX paragraph flows top-to-bottom. PPTX shapes can be anywhere on the slide. The MD flattens this into linear text, losing positional information.
- **Footer/boilerplate repetition.** Every slide has `(c)2026 3Cloud...` and a page number in the MD. These are typically master-slide elements in PPTX, not per-slide content. The pipeline must ignore these during comparison.
- **Multi-shape slides.** A single slide may have a title shape, body shape, image shape, and footer shape. The MD concatenates all text; the update must route text back to the correct shape.

---

## 6. Proposed File Structure

```
src/md2word/
  approaches/
    pandoc.py          # (modify) handle .pptx output extension
    xml_edit.py        # (unchanged) DOCX update mode
    pptx_edit.py       # (NEW) PPTX update mode
  ai/
    client.py          # (unchanged)
    chunk.py           # (unchanged) DOCX chunking
    pptx_chunk.py      # (NEW) PPTX slide extraction
    map.py             # (modify) add slide-aware parsing mode
    edit_plan.py       # (modify) add DrawingML prompt variant
    pptx_edit_plan.py  # (NEW, alternative) separate PPTX edit plan if prompts diverge too much
    summarize.py       # (unchanged)
    conflict.py        # (unchanged, not called for PPTX)
  cli.py               # (modify) route .pptx to pptx_edit approach
  validate.py          # (modify) add python-pptx validation
```

New dependency: `python-pptx` (add to `pyproject.toml`).

---

## 7. Implementation Phases

### Phase 1: Create Mode

- Modify `approaches/pandoc.py` to handle `.pptx` output
- Add slide-boundary preprocessing: convert `<!-- Slide number: N -->` to `---` horizontal rules
- Add `### Notes:` to pandoc `::: notes` block conversion
- Support `--reference-doc` with a PPTX template
- Validate output with python-pptx

**This gets basic MD --> new PPTX working immediately.**

### Phase 2: Update Mode - Text Changes

- `pptx_chunk.py`: extract slides from PPTX zip (title, body, notes per slide)
- `pptx_edit.py`: main update pipeline
- Slide-aware MD parser (split on `<!-- Slide number: -->`)
- Mapping: ordered + title-match alignment (deterministic first, AI fallback)
- Edit plan using Approach B (JSON content description, python-pptx application)
- Text-only updates: title, body paragraphs, bullet levels, bold/italic, notes
- Zip repackage for modified slides

### Phase 3: Update Mode - Structural Changes

- Insert new slides (from MD slides with no PPTX match)
- Delete slides (PPTX slides with no MD match)
- Table updates within slides
- Image handling (new images, image rescue)
- DrawingML XML generation (Approach A) for complex cases
- Full presentation.xml and Content_Types.xml management

### Phase 4: Polish

- Round-trip validation (PPTX --> MD --> PPTX consistency checks)
- Deterministic fallback (no-AI mode) for PPTX
- Unchanged detection (skip slides with no text changes)
- Speaker notes diff and update
- Handle slide master/footer boilerplate filtering

---

## 8. Risks and Open Questions

### High Risk
- **Shape position preservation.** Editing text in a shape can overflow the shape bounds. python-pptx auto-shrinks text, but the visual result may differ from the original.
- **LLM DrawingML generation quality.** DrawingML is more verbose and positional than WordprocessingML. LLMs may struggle with EMU values, placeholder indices, and shape references. Approach B (JSON + python-pptx) mitigates this.
- **Slide insert/delete complexity.** Adding or removing slides requires updating `presentation.xml`, `[Content_Types].xml`, and multiple `.rels` files. python-pptx handles this, but mixing python-pptx with direct zip manipulation may cause conflicts.

### Medium Risk
- **Multi-shape content routing.** When the MD has text that spans multiple shapes (title, body, caption), the pipeline must determine which text goes to which shape. Placeholder type detection (`<p:ph type="title"/>`, `<p:ph idx="1"/>`) helps but isn't always reliable for custom layouts.
- **Table fidelity.** PPTX tables have explicit column widths in EMUs and per-cell formatting. Regenerating from markdown pipe tables loses this formatting.
- **Image alt-text matching.** The Yeti MD uses descriptive alt text (`![A diagram of a business process AI-generated content may be incorrect.]`). Matching these to PPTX image shapes for preservation requires fuzzy text comparison.

### Open Questions
1. **Should insert/delete be supported in v1?** Or only text updates to existing slides? (Recommend: text-only for v1.)
2. **How to handle slides with no title?** Some slides are section dividers or image-only. The Yeti deck has some.
3. **Should the tool support both `<!-- Slide number: N -->` and pandoc-style `---` boundaries?** Or just the comment style since that's what PPTX-to-MD export produces?
4. **python-pptx vs direct XML manipulation?** python-pptx is safer but less granular. Direct XML gives full control but requires managing all the OOXML bookkeeping. Recommendation: python-pptx for structural changes, direct XML for text-level surgery.
5. **Should speaker notes be editable?** The Yeti MD has substantial notes. Updating them is straightforward with python-pptx but adds to the mapping surface area.

---

## 9. Comparison: DOCX vs PPTX Pipeline Complexity

| Concern | DOCX | PPTX | Easier? |
|---------|------|------|---------|
| Chunking | Complex (heading style parsing) | Trivial (per-slide files) | PPTX |
| Section boundaries | Implicit (heading styles) | Explicit (file boundaries) | PPTX |
| Mapping | Fuzzy heading match | Ordered + title match | PPTX |
| Conflict detection | Needed (tracked changes) | Not needed | PPTX |
| Text formatting XML | Simpler (`w:p > w:r > w:t`) | More nested (shape > txBody > `a:p`) | DOCX |
| Table XML | Moderately complex | More complex (graphicFrame + EMUs) | DOCX |
| Image handling | Single rels file | Per-slide rels files | DOCX |
| Slide/section insert | N/A (sections are implicit) | Complex (multi-file update) | DOCX |
| Zip repackage | Replace 1 file | Replace N files | DOCX |
| Layout preservation | Paragraph styles | Shape positions + slide layouts | DOCX |
| Overall | Harder mapping, simpler XML | Simpler mapping, harder XML + structure | ~Even |
