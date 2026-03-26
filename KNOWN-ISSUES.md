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
| Table row reordering | When a row moves position within a table, it may end up at the end instead of its correct spot | Both modes |
| Occasional bullet reordering | The AI may subtly reorder 1-2 bullets in a long list when regenerating a section | AI mode |
| Heading level drift | The AI occasionally renders a heading at the wrong level | AI mode |
| Double space collapse | Multiple consecutive spaces are reduced to one | Both modes |

## Previously resolved

These were significant issues found and fixed during development:

| Issue | What happened | How it was fixed |
|-------|--------------|-----------------|
| Content loss in large sections | A large Markdown section mapped to a small DOCX section, losing embedded sub-sections | Post-mapping re-split detects embedded headings and creates separate mappings |
| Table content loss | AI output dropped complex tables entirely | Original tables are rescued when AI output has none or fewer than 50% of rows |
| Duplicated content across sections | AI regenerated content that already existed in adjacent sections | Three-pass deduplication scans for and removes duplicated paragraphs |
| Table row insertion order | Rows inserted in wrong order when multiple rows added to same table | Switched from index-based to stable element references for insertion anchors |
| Underscore corruption in tables | Identifiers like `GFS_CTS_DEX` became `GFS\CTS\DEX` | Separate table-cell text cleaner that preserves literal underscores |
