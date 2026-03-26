"""
Interpretation step: classify the relationship between an MD file and an
optional base DOCX.

Deterministic — no AI needed:
  - base provided → update
  - no base       → create
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


@dataclass
class MatchResult:
    mode: Literal["create", "update"]
    rationale: str


def interpret(md_text: str, has_base: bool) -> MatchResult:
    """Classify the mode based on whether a base document was provided."""
    if has_base:
        return MatchResult(
            mode="update",
            rationale="Base document provided; updating existing document.",
        )
    return MatchResult(
        mode="create",
        rationale="No base document; creating new document.",
    )
