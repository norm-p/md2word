"""
conflict.py — Detect tracked changes and comments in a base DOCX (update mode).

Fully deterministic: scans DOCX XML for revision marks (w:ins, w:del) and
comment annotations. Tracked changes are always accepted before applying edits.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass
class ConflictReport:
    insertions: int
    deletions: int
    comments: int

    @property
    def has_tracked_changes(self) -> bool:
        return self.insertions > 0 or self.deletions > 0

    @property
    def has_conflicts(self) -> bool:
        return self.has_tracked_changes or self.comments > 0

    @property
    def summary(self) -> str:
        if not self.has_conflicts:
            return "No tracked changes or comments found."
        parts = []
        if self.has_tracked_changes:
            parts.append(f"{self.insertions} insertions, {self.deletions} deletions")
        if self.comments > 0:
            parts.append(f"{self.comments} comments")
        return "Found: " + "; ".join(parts) + "."


def detect_conflicts(docx_path: Path) -> ConflictReport:
    """Scan the DOCX for tracked changes and comments."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")
        comments_xml = (
            zf.read("word/comments.xml")
            if "word/comments.xml" in zf.namelist()
            else None
        )

    root = etree.fromstring(doc_xml)
    insertions = len(root.findall(f".//{{{_W}}}ins"))
    deletions = len(root.findall(f".//{{{_W}}}del"))

    comment_count = 0
    if comments_xml is not None:
        comments_root = etree.fromstring(comments_xml)
        comment_count = len(comments_root.findall(f".//{{{_W}}}comment"))

    return ConflictReport(
        insertions=insertions,
        deletions=deletions,
        comments=comment_count,
    )
