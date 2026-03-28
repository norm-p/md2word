"""
pre_compare.py — Pre-comparison of base DOCX (via MarkItDown) against source MD.

Converts the base DOCX to Markdown, diffs it section-by-section against the
user's updated source MD, and classifies each section so the pipeline can skip
unchanged content, apply deterministic patches for minor edits, and reserve
LLM regeneration for genuinely rewritten sections.
"""

from __future__ import annotations

import unicodedata
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Literal

import click

from .chunk import DocxSection
from .map import SectionMapping, _parse_md_sections, _normalize_for_heading_match

# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

Action = Literal["identical", "minor_edit", "major_change", "new", "deleted"]


@dataclass
class SectionDiff:
    """Result of comparing one source MD section against the base MD."""

    source_heading: str
    source_level: int
    source_content: str
    base_heading: str | None  # None for new sections
    base_content: str | None  # None for new sections
    similarity: float  # 0.0–1.0 content similarity
    action: Action
    heading_renamed: bool = False


# ---------------------------------------------------------------------------
# Normalization — must match across MarkItDown output and source MD
# ---------------------------------------------------------------------------

_EMPHASIS_RE = re.compile(r"\*{1,2}|_{1,2}")
_WHITESPACE_RE = re.compile(r"\s+")


def _norm(text: str) -> str:
    """Normalize text for comparison: strip emphasis, collapse whitespace, lowercase."""
    text = unicodedata.normalize("NFKD", text)
    text = _EMPHASIS_RE.sub("", text)
    text = _WHITESPACE_RE.sub(" ", text).strip().lower()
    return text


def _heading_sim(a: str, b: str) -> float:
    """Similarity between two heading strings after normalization."""
    na, nb = _norm(a), _norm(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb, autojunk=False).ratio()


# ---------------------------------------------------------------------------
# Section alignment
# ---------------------------------------------------------------------------

# Thresholds for classification
_IDENTICAL_THRESHOLD = 0.99
_MINOR_EDIT_THRESHOLD = 0.85
_HEADING_MATCH_THRESHOLD = 0.60


def _align_sections(
    base_sections: list[tuple[str, int, str]],
    source_sections: list[tuple[str, int, str]],
) -> list[SectionDiff]:
    """Align source MD sections to base MD sections by heading + position.

    Uses a greedy forward scan: for each source section, find the best
    unmatched base section by heading similarity, with a positional bias
    towards nearby sections.
    """
    matched_base: set[int] = set()
    diffs: list[SectionDiff] = []

    for si, (s_heading, s_level, s_content) in enumerate(source_sections):
        best_bi: int | None = None
        best_score: float = 0.0

        for bi, (b_heading, b_level, b_content) in enumerate(base_sections):
            if bi in matched_base:
                continue
            hsim = _heading_sim(s_heading, b_heading)
            if hsim < _HEADING_MATCH_THRESHOLD:
                continue
            # Positional bias: prefer nearby sections (within ±5 positions)
            distance = abs(si - bi)
            pos_bonus = max(0, 0.1 - distance * 0.02)
            score = hsim + pos_bonus
            if score > best_score:
                best_score = score
                best_bi = bi

        if best_bi is not None:
            matched_base.add(best_bi)
            b_heading, b_level, b_content = base_sections[best_bi]
            # Content similarity on normalized text
            s_norm = _norm(s_content)
            b_norm = _norm(b_content)
            if not s_norm and not b_norm:
                sim = 1.0
            elif not s_norm or not b_norm:
                sim = 0.0
            else:
                sim = SequenceMatcher(None, s_norm, b_norm, autojunk=False).ratio()

            heading_renamed = _norm(s_heading) != _norm(b_heading)

            if sim >= _IDENTICAL_THRESHOLD:
                action: Action = "identical"
            elif sim >= _MINOR_EDIT_THRESHOLD:
                action = "minor_edit"
            else:
                action = "major_change"

            diffs.append(SectionDiff(
                source_heading=s_heading,
                source_level=s_level,
                source_content=s_content,
                base_heading=b_heading,
                base_content=b_content,
                similarity=sim,
                action=action,
                heading_renamed=heading_renamed,
            ))
        else:
            # No matching base section — new content
            diffs.append(SectionDiff(
                source_heading=s_heading,
                source_level=s_level,
                source_content=s_content,
                base_heading=None,
                base_content=None,
                similarity=0.0,
                action="new",
            ))

    return diffs


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def pre_compare(base_md: str, source_md: str) -> list[SectionDiff]:
    """Diff base MD (from MarkItDown) against source MD, section by section.

    Returns a list of SectionDiff, one per source MD section. Sections that
    exist in the base but not in the source are not included (they will be
    preserved as-is in the DOCX since we never delete content).
    """
    base_sections = _parse_md_sections(base_md)
    source_sections = _parse_md_sections(source_md)

    diffs = _align_sections(base_sections, source_sections)

    return diffs


def pre_compare_overall_similarity(diffs: list[SectionDiff]) -> float:
    """Return the fraction of source sections that are identical or near-identical.

    Used to decide whether the pre-comparison is reliable enough to drive the
    pipeline. If this is below ~0.50, the MarkItDown rendering may differ too
    much from the source MD conventions.
    """
    if not diffs:
        return 0.0
    ok = sum(1 for d in diffs if d.similarity >= _MINOR_EDIT_THRESHOLD)
    return ok / len(diffs)


def map_sections_precompare(
    diffs: list[SectionDiff],
    docx_sections: list[DocxSection],
) -> list[SectionMapping]:
    """Convert pre-comparison diffs into SectionMapping list.

    Bridges: source MD heading → base MD heading (from diff) → DOCX section
    (by heading match against chunk.py output).
    """
    # Build lookup from DOCX heading to DocxSection.
    # Multiple lookups: exact, case-insensitive, normalized.
    docx_by_heading: dict[str, DocxSection] = {s.heading: s for s in docx_sections}
    docx_by_lower: dict[str, DocxSection] = {s.heading.lower(): s for s in docx_sections}
    docx_by_norm: dict[str, DocxSection] = {
        _normalize_for_heading_match(s.heading): s
        for s in docx_sections
        if _normalize_for_heading_match(s.heading)
    }
    claimed: set[str] = set()

    def _find_docx_section(heading: str) -> DocxSection | None:
        """Find DOCX section by heading, trying exact → case-insensitive → normalized."""
        if heading in docx_by_heading and heading not in claimed:
            claimed.add(heading)
            return docx_by_heading[heading]
        low = heading.lower()
        if low in docx_by_lower and docx_by_lower[low].heading not in claimed:
            sec = docx_by_lower[low]
            claimed.add(sec.heading)
            return sec
        norm = _normalize_for_heading_match(heading)
        if norm and norm in docx_by_norm and docx_by_norm[norm].heading not in claimed:
            sec = docx_by_norm[norm]
            claimed.add(sec.heading)
            return sec
        return None

    results: list[SectionMapping] = []

    for diff in diffs:
        # Try to find DOCX section using the BASE heading (more likely to match
        # DOCX heading styles) then fall back to source heading.
        docx_sec: DocxSection | None = None
        if diff.base_heading is not None:
            docx_sec = _find_docx_section(diff.base_heading)
        if docx_sec is None:
            docx_sec = _find_docx_section(diff.source_heading)

        if diff.action == "identical":
            action: Literal["insert", "replace", "unchanged"] = "unchanged"
        elif diff.action == "minor_edit":
            # Deterministic patches (Options B/C/D) handle minor edits.
            action = "unchanged"
        elif diff.action == "major_change":
            action = "replace" if docx_sec is not None else "insert"
        elif diff.action == "new":
            action = "insert"
        else:
            action = "replace"

        results.append(SectionMapping(
            md_heading=diff.source_heading,
            md_heading_level=diff.source_level,
            md_content=diff.source_content,
            docx_section=docx_sec,
            action=action,
        ))

    return results
