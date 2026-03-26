# md2word

Convert Markdown to Word (.docx) -- or **update an existing Word document** from a revised Markdown source while preserving formatting, styles, images, and tables.

# THIS IS AN UPDATED DOCUMENT

Most Markdown-to-Word tools only do one-way conversion. md2word also goes the other direction: give it your edited Markdown alongside the original `.docx`, and it produces an updated document that keeps the original's formatting intact. AI assists with section alignment and edit generation; everything else is deterministic.

**No API key? No problem.** md2word works without any AI provider configured -- it falls back to deterministic-only mode with heading matching, text corrections, bullet patching, and table row updates.

## Modes

| Mode | Trigger | What happens |
|------|---------|--------------|
| **Create** | No base `.docx` provided | Converts Markdown to a new `.docx` via pandoc |
| **Update** | Base `.docx` provided | Preserves base formatting, applies content changes from Markdown |
| **Update (no AI)** | Base `.docx` provided, no API key | Deterministic heading matching + targeted patches (bullets, text, table rows) |

## Setup

**Requirements:** Python 3.12+, [uv](https://docs.astral.sh/uv/)

```bash
git clone https://github.com/norm-p/md2word.git && cd md2word
uv sync                    # core dependencies (deterministic mode)
uv sync --extra anthropic  # + Anthropic AI provider (optional)
```

Pandoc is bundled as a portable binary (`tools/pandoc.exe`) -- no system install needed.

### Optional: configure an AI provider

Copy `.env.example` to `.env` and set `AI_PROVIDER`, `AI_MODEL`, and the matching API key. This unlocks AI-augmented section mapping, OOXML generation, and change summaries in update mode.

Without a configured provider (or without the AI SDK installed), update mode runs in deterministic-only mode (see below).

## Quick start

Try it on this README:

```bash
uv run md2word README.md
uv run md2word samples/README_updated.md README.docx --accept-changes
```

The first command creates `README.docx`. The second updates it from the revised Markdown, producing `README_v001.docx` with the changes applied.

## Usage

```bash
uv run md2word document.md
uv run md2word document.md --ref-doc template.docx
uv run md2word revised.md original.docx
uv run md2word revised.md original.docx -o updated.docx
uv run md2word -i revised.md -b original.docx -o updated.docx
uv run md2word revised.md original.docx --accept-changes
uv run md2word document.md original.docx -v
```

### Options

| Flag | Description |
|------|-------------|
| `-i, --input PATH` | Markdown input file (or first positional arg) |
| `-b, --base PATH` | Existing `.docx` to update (or second positional arg) |
| `-o, --output PATH` | Output file path (default: versioned from base, or input stem + `.docx`) |
| `--ref-doc PATH` | Reference `.docx` for pandoc styling (create mode only) |
| `--toc` | Inject table of contents placeholder (create mode only) |
| `--accept-changes` | Skip change summary and confirmation (update mode) |
| `-v, --verbose` | Show step-by-step progress |

### Output versioning

In update mode, output auto-versions from the base filename:

```
report.docx      ->  report_v001.docx
report_v001.docx ->  report_v002.docx
```

Existing files are never overwritten.

## How update mode works

The update pipeline is primarily deterministic. AI is used only for three steps that require semantic judgment -- everything else is deterministic Python.

```
INPUT.md + BASE.docx
  1. Scan for tracked changes/comments      (deterministic)
  2. Chunk DOCX into heading-delimited sections  (deterministic)
  3. Map MD sections to DOCX sections        (AI or deterministic)
  4. Detect unchanged sections by text similarity  (deterministic)
  5. Generate OOXML edit fragments           (AI -- skipped without API key)
  6. Summarize changes for confirmation      (AI -- skipped without API key)
  7. Apply edits via XML surgery             (deterministic)
     - Preserve original heading styles
     - Rescue images and tables the AI can't regenerate
     - Deduplicate LLM-generated content
     - Patch unchanged sections (bullets, text, table rows)
     - Enforce source bullet order
  8. Repackage .docx preserving zip metadata (deterministic)
  9. Validate output is not corrupt          (deterministic)
```

### Deterministic-only mode (no API key)

When no AI provider is configured, update mode:

1. Matches sections by heading text (exact, case-insensitive, and normalized fuzzy matching)
2. Detects unchanged sections via text similarity (>= 90% threshold)
3. Applies targeted patches to unchanged sections:
   - Inserts new bullet points
   - Corrects minor text differences in-place
   - Updates, inserts, and removes table rows
4. Preserves all original formatting, images, and styles

This is useful for documents with small, targeted edits where the Markdown changes are incremental.

## Supported AI providers

Set `AI_PROVIDER` and `AI_MODEL` in `.env`:

| Provider | Value | Required env vars |
|----------|-------|-------------------|
| Anthropic | `anthropic` | `ANTHROPIC_API_KEY` |
| Azure AI Foundry | `azure` | `AZURE_ANTHROPIC_KEY`, `AZURE_ANTHROPIC_ENDPOINT` |
| AWS Bedrock | `bedrock` | `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY` |
| Google Vertex AI | `vertex` | `VERTEX_PROJECT_ID`, `VERTEX_REGION` |
| OpenAI | `openai` | `OPENAI_API_KEY` |
| Azure OpenAI | `azure-openai` | `AZURE_OPENAI_KEY`, `AZURE_OPENAI_ENDPOINT` |

Some providers require optional dependencies:

```bash
uv sync --extra bedrock   # AWS Bedrock
uv sync --extra vertex    # Google Vertex AI
uv sync --extra openai    # OpenAI or Azure OpenAI
```

## Architecture

```
src/md2word/
  cli.py                  # Entry point, mode detection, output versioning
  validate.py             # Output validation (python-docx open check)
  ai/
    client.py             # Provider-aware LLM client
    chunk.py              # DOCX section chunking (deterministic)
    map.py                # Section alignment (AI + deterministic fallback)
    edit_plan.py          # OOXML fragment generation (AI, batched)
    conflict.py           # Tracked-change detection (deterministic)
    summarize.py          # Change summary for confirmation (AI)
    interpret.py          # Mode detection (deterministic)
  approaches/
    pandoc.py             # Create mode (pandoc conversion)
    xml_edit.py           # Update mode (XML surgery pipeline)
```

**Design principle:** AI is used only where semantic judgment is needed (section alignment, OOXML generation, change summarization). Everything else -- mode detection, conflict scanning, text similarity, XML manipulation, validation -- is deterministic Python with [lxml](https://lxml.de/).

## Contributing

Contributions are welcome. The codebase follows a clear separation: AI logic lives in `src/md2word/ai/`, deterministic approaches in `src/md2word/approaches/`. See `CLAUDE.md` for detailed architecture notes and `WORKFLOW.md` for the full pipeline documentation with diagrams.

## License

MIT
