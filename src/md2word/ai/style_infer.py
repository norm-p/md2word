"""
style_infer.py — Infer Word styles for ambiguous MD elements (create mode).

Pandoc handles common mappings (# → Heading 1, etc.) but can be ambiguous for
custom styles. This module asks the LLM to map uncertain elements to the styles
available in the reference template.
"""

from __future__ import annotations

import json

from .client import LLMClient, parse_llm_json

_SYSTEM = """\
You are a Word style mapper. You will receive:

1. The text of a Markdown document.
2. A list of available Word styles from a reference template.

Your job is to identify Markdown elements that should use a non-default Word
style, and map them. Respond with a JSON object — no markdown fences.

The keys are Markdown element patterns (e.g. "## Introduction", "**Note:**")
and the values are the exact style names from the available list.

Rules:
- Only include elements that need NON-DEFAULT style treatment.
- Standard mappings (# = Heading 1, ## = Heading 2, etc.) should NOT be listed
  unless the template uses non-standard naming.
- Focus on: blockquotes, admonitions, custom callouts, table captions, and any
  other elements where pandoc's default mapping would be wrong.
- If no custom mappings are needed, return an empty object {}.
"""


def infer_styles(
    client: LLMClient,
    md_text: str,
    available_styles: list[str],
) -> dict[str, str]:
    """Return a mapping of MD element -> Word style name.

    Only elements that require non-default style treatment are included.
    """
    user_content = (
        f"Markdown document:\n\n{md_text}\n\n---\n\n"
        f"Available Word styles:\n{json.dumps(available_styles)}"
    )

    raw = client.complete(
        system=_SYSTEM,
        messages=[{"role": "user", "content": user_content}],
    )

    try:
        return parse_llm_json(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"LLM returned invalid JSON in style_infer:\n{raw}") from exc
