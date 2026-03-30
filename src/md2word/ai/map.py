"""
map.py — Map MD sections to DOCX sections (update mode).

Uses the LLM to align headings and content between the incoming Markdown and
the existing DOCX sections produced by chunk.py. Returns a SectionMapping list
that drives edit_plan.py.

Large documents (>30 MD sections) are processed in batches. All DOCX headings
are sent to every batch so the LLM can find the best match for each MD section.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Literal

import click

from .. import output as out
from .chunk import DocxSection
from .client import LLMClient, parse_llm_json

_MAP_BATCH_SIZE = 30  # MD sections per LLM call

# Regex to strip XML tags — used to estimate plain-text length from OOXML fragments
_TAG_RE = re.compile(r"<[^>]+>")

# Patterns for normalising MD lines before heading comparison (R2-A)
_BOLD_MARKER_RE = re.compile(r"^\*{1,2}|\*{1,2}$")
_LIST_PREFIX_RE = re.compile(r"^(?:\*\s+|-\s+|\d+\.\s+)+")


def _normalize_for_heading_match(stripped: str) -> str:
    """Strip bold markers and list prefixes from a line for loose heading comparison.

    Used as a fallback when exact case-insensitive matching fails.
    E.g. ``**Schedule**`` → ``schedule``, ``* 1. Schedule`` → ``schedule``.
    """
    s = _BOLD_MARKER_RE.sub("", stripped).strip()
    s = _LIST_PREFIX_RE.sub("", s).strip()
    s = _BOLD_MARKER_RE.sub("", s).strip()
    return s.lower()


@dataclass
class SectionMapping:
    md_heading: str
    md_heading_level: int   # 1-9 detected from MD; 0 for preamble
    md_content: str
    docx_section: DocxSection | None
    action: Literal["insert", "replace", "unchanged"]


_SYSTEM = """\
You are a document section mapper. You will receive:

1. A list of Markdown sections (heading + content preview).
2. A list of existing Word document sections (heading only).

Your job is to align each Markdown section to its best-matching Word section,
or mark it as new. Respond with a JSON array — no markdown fences, no extra text.

Each element:
{
  "md_heading": "<exact MD heading text>",
  "docx_heading": "<exact DOCX heading text>" | null,
  "action": "replace" | "insert" | "unchanged"
}

Rules:
- "replace"   — MD section updates/replaces an existing DOCX section.
- "insert"    — MD section is new, no DOCX equivalent. Set docx_heading to null.
- "unchanged" — MD section content is semantically identical to the DOCX section;
                no edit needed.
- Match by semantic meaning, not just exact heading text (e.g. "Intro" matches
  "Introduction").
- Every MD section must appear exactly once in the output.
- A DOCX section can be matched to at most one MD section.
"""


###############################################################################
# Boilerplate separator detection (dynamic) — must use the same criteria as
# chunk.py so both parsers split at matching boundaries.
# In Markdown, bold all-caps lines (**TEXT**) correspond to bold all-caps
# paragraphs in the DOCX.
###############################################################################

_SEPARATOR_MIN_LEN = 8
_SEPARATOR_MAX_LEN = 80


def _boilerplate_section_name_md(line: str) -> str | None:
    """Return a synthetic section heading if a Markdown line is a boilerplate separator.

    Matches bold all-caps lines like ``**SIGNATURE PAGE FOLLOWS**``.
    Returns a normalized heading like ``(signature page follows)`` that
    matches the synthetic heading chunk.py generates from the DOCX side.
    """
    stripped = line.strip()
    # Must be bold-wrapped: **TEXT**
    if not (stripped.startswith("**") and stripped.endswith("**")):
        return None
    inner = stripped[2:-2].strip()
    if not (_SEPARATOR_MIN_LEN <= len(inner) <= _SEPARATOR_MAX_LEN):
        return None
    alpha = [c for c in inner if c.isalpha()]
    if len(alpha) < 4 or not all(c.isupper() for c in alpha):
        return None
    return f"({inner.lower()})"


###############################################################################
# Heading patterns
###############################################################################

# Standard ATX headings:  # Title  /  ## Section  etc.
_ATX_RE = re.compile(r"^(#{1,6})\s+(.+?)(?:\s+#+)?\s*$")

# Numbered-list headings exported from Word, e.g.:
#   "1. Executive Summary"            → level 1  (no indent, no bullet)
#   "* 1. Preflight"                  → level 2  (bullet + number)
#   "  1. Takeoff"                    → level 2  (1-2 space indent + number)
# These appear when Word autonumber styles are exported to Markdown.
# NOTE: Lines starting with bold (**text**) are list items, not headings.
# Lines with 3+ space indent are sub-list items, not structural headings.
_NUMBERED_H1 = re.compile(r"^\d+\.\s+(.+)$")
_NUMBERED_H2 = re.compile(r"^(?:\*\s+|\s{1,2})\d+\.\s+(.+)$")


def _detect_heading(line: str, use_numbered: bool) -> tuple[int, str] | None:
    """Return (level, text) if line is a heading, else None."""
    m = _ATX_RE.match(line)
    if m:
        return len(m.group(1)), m.group(2).strip()
    if use_numbered:
        m = _NUMBERED_H1.match(line)
        if m:
            text = m.group(1).strip()
            # Bold-prefixed numbered items (1. **text**) are list items, not headings
            if text.startswith("**"):
                return None
            return 1, text
        m = _NUMBERED_H2.match(line)
        if m:
            return 2, m.group(1).strip()
    return None


def _parse_md_sections(md_text: str) -> list[tuple[str, int, str]]:
    """Split markdown into (heading, level, content) triples.

    Returns a list of triples. Content before the first heading gets heading
    "(preamble)" at level 0. Handles trailing # markers on ATX headings.

    When no ATX headings are found (e.g., a DOCX export where numbered list
    items represent section headers), numbered-list heading detection is
    enabled as a fallback.
    """
    lines = md_text.split("\n")

    # Use numbered headings unless ATX headings are the dominant format.
    # A document with only 1-2 ATX headings among many numbered-list headings
    # (e.g., a DOCX export) should still use numbered heading detection.
    atx_count = sum(1 for l in lines if _ATX_RE.match(l))
    use_numbered = atx_count < 3

    sections: list[tuple[str, int, str]] = []
    current_heading = "(preamble)"
    current_level = 0
    current_lines: list[str] = []

    for line in lines:
        hit = _detect_heading(line, use_numbered)
        if hit is not None:
            if current_lines or current_heading == "(preamble)":
                sections.append((current_heading, current_level, "\n".join(current_lines).strip()))
            current_level, current_heading = hit[0], hit[1]
            current_lines = []
            continue

        # Boilerplate markers (e.g. "**SIGNATURE PAGE FOLLOWS**") create a
        # synthetic section break matching the same split in chunk.py.
        marker = _boilerplate_section_name_md(line)
        if marker is not None:
            if current_lines or current_heading == "(preamble)":
                sections.append((current_heading, current_level, "\n".join(current_lines).strip()))
            current_heading = marker
            current_level = 0
            current_lines = [line]
            continue

        current_lines.append(line)

    sections.append((current_heading, current_level, "\n".join(current_lines).strip()))

    # Drop empty preamble
    if sections and sections[0][0] == "(preamble)" and not sections[0][2]:
        sections = sections[1:]

    return sections


def _call_map_batch(
    client: LLMClient,
    md_batch: list[tuple[str, str]],
    docx_list: list[dict],
) -> list[dict]:
    """Call the LLM to map a batch of MD sections against the full DOCX section list."""
    md_list = [{"heading": h, "content_preview": c[:200]} for h, c in md_batch]  # pairs only
    user_content = (
        f"Markdown sections:\n{json.dumps(md_list, indent=2)}\n\n"
        f"Word document sections:\n{json.dumps(docx_list, indent=2)}"
    )

    raw = client.complete(system=_SYSTEM, messages=[{"role": "user", "content": user_content}])

    try:
        return parse_llm_json(raw)
    except json.JSONDecodeError as exc:
        # Retry with error feedback
        retry_messages = [
            {"role": "user", "content": user_content},
            {"role": "assistant", "content": raw},
            {
                "role": "user",
                "content": (
                    f"Your response was not valid JSON: {exc}. "
                    "Return only a valid JSON array, no extra text or markdown fences."
                ),
            },
        ]
        raw2 = client.complete(system=_SYSTEM, messages=retry_messages)
        return parse_llm_json(raw2)


def _resplit_large_sections(
    results: list[SectionMapping],
    docx_sections: list[DocxSection],
) -> list[SectionMapping]:
    """Post-mapping re-split for large-MD/small-DOCX mismatches (Issue 2).

    After the LLM mapper runs, a large MD section may map to a small DOCX
    section with low similarity because the MD contains embedded sub-section
    text whose headings match unmatched DOCX sections.  The DOCX heading list
    is ground truth — when an unmatched DOCX heading appears verbatim as a
    full line inside the large MD section, we split at that line and create a
    direct SectionMapping for each sub-section, bypassing the LLM entirely.

    Trigger conditions (all must be met):
    - Mapping action is "replace" with a matched DOCX section
    - MD content is ≥ 2,000 characters
    - MD content length is ≥ 2.5× the estimated DOCX section text length
    - At least one unmatched DOCX heading appears as a full (stripped) line
      in the MD content (case-insensitive)
    """
    # DOCX headings already claimed by the mapper
    matched_docx_headings: set[str] = {
        m.docx_section.heading
        for m in results
        if m.docx_section is not None
    }
    unmatched_docx = [s for s in docx_sections if s.heading not in matched_docx_headings]
    if not unmatched_docx:
        return results

    # Case-insensitive lookup: heading.lower() → DocxSection
    unmatched_by_lower: dict[str, DocxSection] = {
        s.heading.lower(): s for s in unmatched_docx
    }

    new_results: list[SectionMapping] = []

    for mapping in results:
        if (
            mapping.action != "replace"
            or mapping.docx_section is None
            or len(mapping.md_content) < 2000
        ):
            new_results.append(mapping)
            continue

        # Estimate DOCX section text length by stripping XML tags
        docx_text = _TAG_RE.sub(" ", mapping.docx_section.xml_fragment)
        docx_text = re.sub(r"\s+", " ", docx_text).strip()

        # Only trigger on significant size mismatch
        if len(mapping.md_content) < len(docx_text) * 2.5:
            new_results.append(mapping)
            continue

        # Scan MD content for unmatched DOCX headings appearing as full lines.
        # Primary: exact match after stripping whitespace (case-insensitive).
        # Fallback: normalised match stripping bold markers and list prefixes
        # (catches e.g. "**Schedule**" matching DOCX heading "Schedule").
        lines = mapping.md_content.split("\n")
        split_points: list[tuple[int, DocxSection]] = []  # (line_index, docx_section)
        for i, line in enumerate(lines):
            stripped = line.strip()
            if not stripped:
                continue
            lower = stripped.lower()
            if lower in unmatched_by_lower:
                split_points.append((i, unmatched_by_lower[lower]))
            else:
                norm = _normalize_for_heading_match(stripped)
                if norm and norm != lower and norm in unmatched_by_lower:
                    split_points.append((i, unmatched_by_lower[norm]))
                elif len(lower) >= 10:
                    # Prefix match: DOCX heading may concatenate heading + subtitle
                    # (e.g. "EXHIBIT E TO SOW# 223571FEES" vs MD line "EXHIBIT E TO SOW# 223571")
                    for uh_lower, docx_sec in unmatched_by_lower.items():
                        if uh_lower.startswith(lower):
                            split_points.append((i, docx_sec))
                            break

        if not split_points:
            new_results.append(mapping)
            continue

        out.detail(
            f"Re-splitting '{mapping.md_heading}': "
            f"found {len(split_points)} embedded DOCX heading(s) → "
            f"{len(split_points) + 1} synthetic sections"
        )

        # Intro: text before the first split point — keep under the original mapping
        first_idx = split_points[0][0]
        intro_content = "\n".join(lines[:first_idx]).strip()
        new_results.append(SectionMapping(
            md_heading=mapping.md_heading,
            md_heading_level=mapping.md_heading_level,
            md_content=intro_content,
            docx_section=mapping.docx_section,
            action=mapping.action,
        ))

        # Each split point becomes a new replace mapping
        for k, (line_idx, docx_sec) in enumerate(split_points):
            next_idx = split_points[k + 1][0] if k + 1 < len(split_points) else len(lines)
            body_content = "\n".join(lines[line_idx + 1:next_idx]).strip()
            new_results.append(SectionMapping(
                md_heading=docx_sec.heading,
                md_heading_level=docx_sec.heading_level,
                md_content=body_content,
                docx_section=docx_sec,
                action="replace",
            ))
            # Mark this DOCX section as now matched so it can't be reused
            unmatched_by_lower.pop(docx_sec.heading.lower(), None)

    return new_results


def map_sections_deterministic(
    md_text: str,
    docx_sections: list[DocxSection],
) -> list[SectionMapping]:
    """Align MD sections with DOCX sections using deterministic heading matching.

    Used when no LLM API key is available. Matches by: exact text, then
    case-insensitive, then normalized fuzzy (stripping bold/list prefixes).
    Unmatched MD sections become inserts. All matched sections are marked
    "replace" — downstream unchanged detection will downgrade to "unchanged"
    where text is equivalent.
    """
    md_sections = _parse_md_sections(md_text)
    docx_by_heading: dict[str, DocxSection] = {s.heading: s for s in docx_sections}
    docx_by_lower: dict[str, DocxSection] = {s.heading.lower(): s for s in docx_sections}
    docx_by_norm: dict[str, DocxSection] = {
        _normalize_for_heading_match(s.heading): s for s in docx_sections
        if _normalize_for_heading_match(s.heading)
    }
    claimed: set[str] = set()

    results: list[SectionMapping] = []
    for heading, level, content in md_sections:
        docx_sec: DocxSection | None = None

        # Try exact match
        if heading in docx_by_heading and heading not in claimed:
            docx_sec = docx_by_heading[heading]
        # Try case-insensitive
        elif heading.lower() in docx_by_lower and docx_by_lower[heading.lower()].heading not in claimed:
            docx_sec = docx_by_lower[heading.lower()]
        # Try normalized (strip bold, list prefixes)
        else:
            norm = _normalize_for_heading_match(heading)
            if norm and norm in docx_by_norm and docx_by_norm[norm].heading not in claimed:
                docx_sec = docx_by_norm[norm]

        if docx_sec is not None:
            claimed.add(docx_sec.heading)
            action = "replace"
        else:
            action = "insert"

        results.append(SectionMapping(
            md_heading=heading,
            md_heading_level=level,
            md_content=content,
            docx_section=docx_sec,
            action=action,
        ))

    results = _resplit_large_sections(results, docx_sections)
    return results


def map_sections(
    client: LLMClient,
    md_text: str,
    docx_sections: list[DocxSection],
) -> list[SectionMapping]:
    """Ask the LLM to align MD sections with DOCX sections."""
    md_sections = _parse_md_sections(md_text)
    docx_list = [{"heading": s.heading} for s in docx_sections]
    docx_by_heading = {s.heading: s for s in docx_sections}
    md_content_by_heading = {h: c for h, _lvl, c in md_sections}
    md_levels_by_heading = {h: lvl for h, lvl, _c in md_sections}
    # Pairs (heading, content) for the LLM — level is not needed by the mapper
    md_pairs = [(h, c) for h, _lvl, c in md_sections]

    # Batch if the document has many sections
    mappings_data: list[dict] = []
    if len(md_pairs) <= _MAP_BATCH_SIZE:
        mappings_data = _call_map_batch(client, md_pairs, docx_list)
    else:
        total = (len(md_pairs) + _MAP_BATCH_SIZE - 1) // _MAP_BATCH_SIZE
        for i in range(total):
            batch = md_pairs[i * _MAP_BATCH_SIZE:(i + 1) * _MAP_BATCH_SIZE]
            out.detail(f"Mapping sections: batch {i + 1}/{total}")
            mappings_data.extend(_call_map_batch(client, batch, docx_list))

    results: list[SectionMapping] = []
    for item in mappings_data:
        md_heading = item["md_heading"]
        docx_heading = item.get("docx_heading")
        action = item["action"]

        docx_sec = docx_by_heading.get(docx_heading) if docx_heading else None
        md_content = md_content_by_heading.get(md_heading, "")
        md_level = md_levels_by_heading.get(md_heading, 1)

        results.append(SectionMapping(
            md_heading=md_heading,
            md_heading_level=md_level,
            md_content=md_content,
            docx_section=docx_sec,
            action=action,
        ))

    results = _resplit_large_sections(results, docx_sections)
    return results
