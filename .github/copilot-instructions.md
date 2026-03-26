# Project Guidelines

## Code Style
- Use Python 3.12+ features and keep `from __future__ import annotations` where present.
- Prefer explicit type hints and `@dataclass` for structured payloads/results.
- Keep module-level docstrings that explain purpose and pipeline role.
- Favor clear, deterministic error handling with user-facing CLI messages via `click.echo(..., err=True)` and non-zero exits for failures.

## Architecture
- Main orchestration lives in `src/md2word/cli.py`.
- Processing pipeline: interpret (deterministic) -> approach selection (deterministic) -> dispatch -> output validation.
- AI reasoning and provider logic belong in `src/md2word/ai/` modules only.
- Deterministic conversion implementations go in `src/md2word/approaches/`.
- AI is optional: when no API key is configured, update mode uses deterministic heading matching and targeted patching.

## Build and Test
- Install dependencies: `uv sync`
- Run CLI: `uv run md2word INPUT.md [BASE.docx] [-o OUT.docx] [--ref-doc template.docx] [--toc] [-v]`
- Run tests: `uv run pytest`
- Configure `.env` from `.env.example` for AI provider credentials (optional).
- `PYPANDOC_PANDOC` points to the pandoc binary if not on `PATH` (e.g., `tools/pandoc.exe`).

## Conventions
- Keep all LLM exchanges JSON-structured; avoid markdown-fenced outputs in prompts/parsing.
- For Word OOXML operations, keep namespace constants centralized and reused.
- Put provider-specific configuration only in `src/md2word/ai/client.py`.
- Keep approach modules focused: `pandoc.py` for create, `xml_edit.py` for update.
- Validate produced outputs with `validate_output` before reporting success.
- Never modify source MD or base DOCX files; all writes go to output files.

## Common Pitfalls
- Do not assume system pandoc is installed; respect `PYPANDOC_PANDOC`.
- Do not assume an AI provider is configured; `get_client_or_none()` may return `None`.
- If adding new AI providers, update provider selection and env variable handling consistently.
- Use `lxml` for all XML operations (namespace-safe parsing and serialization).
