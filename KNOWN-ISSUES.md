# Known Issues & Limitations

## What to expect

md2word handles most Markdown-to-Word update scenarios well, but there are
edge cases to be aware of depending on which mode you're using.

## Without an AI provider (deterministic mode)

Deterministic mode works great for incremental edits -- correcting text,
adding bullets, updating table rows. These are its limitations:

- **New sections are skipped.** If your Markdown adds an entirely new heading
  that doesn't exist in the DOCX, it won't be inserted. The tool can only
  patch sections it can match by heading text.
- **Large rewrites are left unchanged.** Sections where less than 90% of the
  text matches are preserved as-is. The tool won't attempt to rewrite content
  it can't confidently map.
- **Heading matching is literal.** Matches are exact or case-insensitive
  (with bold/list prefix stripping). A heading like "Intro" won't match
  "Introduction" -- that requires AI.

## With an AI provider

AI mode handles section alignment, content rewrites, and new section insertion.
These are its known quirks:

- **Minor formatting drift.** The AI may apply bold slightly differently than
  the source (e.g., bolding a whole table cell instead of just a name), shift
  bold boundaries around punctuation, or normalize curly quotes to straight
  quotes. Content accuracy is not affected.
- **Complex tables are preserved, not regenerated.** Tables with merged cells,
  custom borders, or precise column widths are carried forward from the
  original DOCX. Content changes within them are applied via targeted row
  patching.
- **Images are carried forward.** Embedded images from the original DOCX are
  preserved. New images referenced in the Markdown are not added.

## Open issues

| Issue | Description | Affects |
|-------|-------------|---------|
| Sub-heading duplication | The AI may emit a sub-heading as a plain paragraph, duplicating text already present from the original section. Most within-section duplicates are removed automatically; a cross-section duplicate may remain (1 extra occurrence). | AI mode |
| Occasional bullet reordering | The AI may subtly reorder 1-2 bullets in a long list when regenerating a section | AI mode |
| Heading level drift | The AI occasionally renders a heading at the wrong level | AI mode |
| Double space collapse | Multiple consecutive spaces are reduced to one | Both modes |

## Recent fixes

Fixes are listed newest-first with the review that introduced them.

### R25 -- Stale table values in AI and unmatched sections

- **Stale table cell values corrected in AI-regenerated sections.** When the AI
  regenerates a section containing a table, deterministic table row patching now
  runs as a second pass to correct any stale values the AI carried over from the
  original DOCX. Previously, only unchanged sections received table patching.
- **Unmatched DOCX sections receive table patching.** DOCX sections whose
  headings do not match any Markdown section (and therefore pass through
  unchanged) now receive deterministic table row patching using the full
  Markdown text. This corrects stale values in sections that were previously
  unreachable by the patching infrastructure.

### R24 -- Duplicate paragraph removal and empty heading cleanup

- **Consecutive duplicate paragraphs removed.** After AI replacement, adjacent
  paragraphs with identical normalized text are detected and the duplicate is
  removed. This eliminates within-section sub-heading duplication.
- **Empty heading paragraphs removed.** A post-edit scan removes heading-styled
  paragraphs with no visible text, catching empty headings that the section-level
  filter could not reach.

### R23 -- Bold-aware table patching and SDT handling

- **Table cell bold boundaries preserved.** When a Markdown table cell contains
  `**bold text**` followed by plain text, the output DOCX now creates separate
  runs with correct bold/plain formatting. Previously, the entire cell was
  rendered bold because formatting markers were stripped before the cell-writing
  code could use them. Affects milestone tables, exhibit tables, and any table
  with mixed bold/plain cells.
- **Structured document tag (SDT) paragraphs skipped in text correction.**
  Paragraphs containing OOXML content controls (`w:sdt`) are no longer processed
  by the text-correction pass. These controls hold text outside normal runs, and
  attempting to correct them caused duplicate text in the output.
- **Heading dedup tightened.** The duplicate-heading removal pass now only checks
  the first two body paragraphs after a section heading, preventing false
  removals deeper in the section content.

### R22 -- Pandoc HTML validation and prompt guidance

- **Pandoc HTML validation tool** (`tools/validate_html.py`). Cross-references
  DOCX output against an independent HTML conversion to distinguish real defects
  from round-trip Markdown artifacts.
- **LLM prompt guidance** added to `edit_plan.py`: bold scope, heading dedup,
  content order, and source priority rules. These serve as soft guidance and do
  not replace the deterministic fixes above.

### R21 -- Table cell cleaning and row stability

- **Table cell markdown unescaping.** Cells containing underscores (e.g.,
  `GFS_CTS_DEX_IVR_01`) are no longer corrupted by overly aggressive markdown
  stripping. A dedicated table-cell cleaner preserves underscores while stripping
  other markdown emphasis.
- **Row insertion anchor stability.** Inserted table rows now use stable element
  references instead of index-based lookups, preventing row order swaps when
  multiple rows are inserted in the same table.
- **Cross-section heading dedup.** Duplicate headings that span section
  boundaries (e.g., a client reference heading appearing in both an edited and
  non-edited section) are detected and removed.

### R20 -- Bullet ordering, content dedup, and text correction

- **Bullet order enforcement.** After patching, bullet lists are reordered to
  match the source Markdown within each contiguous group, preventing cross-group
  shuffling.
- **Spurious LLM table removal.** Tables generated by the AI that don't
  correspond to any table in the original section are removed automatically.
- **Unicode-normalized dedup.** Content deduplication now normalizes smart quotes
  and other Unicode variants before comparison, catching cases the AI normalizes
  differently.
- **Post-LLM text correction.** AI-edited sections receive a second pass of
  deterministic text correction (at a tighter 0.90 threshold) to catch words or
  phrases the AI dropped.
- **Insert boundary fix.** Text insertions at exact run boundaries in the
  correction pass no longer silently drop content.