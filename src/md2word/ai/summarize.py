"""
summarize.py — AI-generated summary of planned changes (update mode).

After the edit plan is built, this module asks the LLM to produce a
human-readable summary of all changes. The CLI presents this summary
and asks for y/n confirmation before applying.

Bypassed by --accept-changes.
"""

from __future__ import annotations

import json
import re

from .client import LLMClient
from .edit_plan import Edit
from .map import SectionMapping


_SYSTEM = """\
You are a document change summarizer. You will receive:

1. A list of section mappings (how Markdown sections align to Word sections).
2. A list of planned edits (insert/replace/delete operations).

Produce a clear, concise summary of what will change in the Word document.
Format as a bulleted list. Be specific about what content is being added,
changed, or removed. Keep each bullet to one sentence.

Respond with plain text only — no JSON, no markdown fences.
"""


def summarize_changes(
    client: LLMClient,
    mapping: list[SectionMapping],
    edits: list[Edit],
) -> str:
    """Return a human-readable summary of the planned changes."""
    mapping_data = []
    for m in mapping:
        entry = {
            "md_heading": m.md_heading,
            "action": m.action,
            "docx_heading": m.docx_section.heading if m.docx_section else None,
            "content_preview": m.md_content[:300] if m.md_content else "",
        }
        mapping_data.append(entry)

    edits_data = [
        {"kind": e.kind, "target_heading": e.target_heading}
        for e in edits
    ]

    user_content = (
        f"Section mappings:\n{json.dumps(mapping_data, indent=2)}\n\n"
        f"Planned edits:\n{json.dumps(edits_data, indent=2)}"
    )

    result = client.complete(system=_SYSTEM, messages=[{"role": "user", "content": user_content}])

    # Strip markdown fences if the model includes them despite instructions
    result = re.sub(r"^```\w*\s*\n?", "", result.strip(), count=1)
    result = re.sub(r"\n?```\s*$", "", result, count=1)
    return result.strip()
