#!/usr/bin/env python3
"""Pandoc HTML validation for md2word output DOCX files.

Converts DOCX -> HTML via pandoc and runs automated checks for known issues
that may be round-trip MD artifacts vs. real DOCX defects.

Usage:
    uv run python tools/validate_html.py OUTPUT.docx [--source SOURCE.md]
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pypandoc
from lxml import html as lhtml
from lxml.html import HtmlElement


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Finding:
    issue_id: str
    title: str
    status: str  # PASS, FAIL, INCONCLUSIVE, NOT_APPLICABLE, NOT_REPRODUCED
    details: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Conversion
# ---------------------------------------------------------------------------

def convert_docx_to_html(docx_path: Path) -> str:
    """Convert DOCX to HTML5 via pandoc. Returns HTML string."""
    return pypandoc.convert_file(
        str(docx_path),
        "html5",
        extra_args=["--standalone", "--wrap=none"],
    )


def save_html(html_content: str, docx_path: Path) -> Path:
    """Save HTML alongside the DOCX and return the path."""
    html_path = docx_path.with_suffix(".html")
    html_path.write_text(html_content, encoding="utf-8")
    return html_path


# ---------------------------------------------------------------------------
# HTML parsing helpers
# ---------------------------------------------------------------------------

def _heading_level(tag: str) -> int | None:
    """Return heading level (1-6) or None."""
    if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
        return int(tag[1])
    return None


def find_table_after_element(tree: HtmlElement, marker_text: str) -> HtmlElement | None:
    """Find the first <table> after any element whose text contains *marker_text* (case-insensitive).

    Searches headings first, then falls back to paragraphs — pandoc sometimes renders
    DOCX heading-styled text as <p> rather than <hN>.
    """
    target = marker_text.lower()
    # Try headings first, then paragraphs
    for tag_filter in (lambda el: _heading_level(el.tag) is not None, lambda el: el.tag == "p"):
        for el in tree.iter():
            if tag_filter(el) and target in (el.text_content() or "").lower():
                for sibling in el.itersiblings():
                    if sibling.tag == "table":
                        return sibling
                    for child in sibling.iter("table"):
                        return child
    return None


def find_section_elements(tree: HtmlElement, heading_text: str) -> list[HtmlElement]:
    """Return all elements between a heading matching *heading_text* and the next same-or-higher heading."""
    target = heading_text.lower()
    collecting = False
    collected: list[HtmlElement] = []
    heading_level = None

    for el in tree.iter():
        lvl = _heading_level(el.tag)
        if lvl is not None and target in (el.text_content() or "").lower():
            collecting = True
            heading_level = lvl
            continue
        if collecting:
            if lvl is not None and lvl <= heading_level:
                break
            collected.append(el)
    return collected


def get_bold_items_in_element(el: HtmlElement) -> list[dict]:
    """Return list of {bold_text, after_text} for each <strong> inside el."""
    items = []
    for strong in el.iter("strong"):
        bold_text = strong.text_content()
        tail = strong.tail or ""
        items.append({"bold_text": bold_text, "after_text": tail.strip()[:80]})
    return items


# ---------------------------------------------------------------------------
# Issue checks
# ---------------------------------------------------------------------------

def check_f5_milestones_bold(tree: HtmlElement, source_md: str | None) -> Finding:
    """F5: Milestones table bold scope — is bold limited to milestone name or covers full cell?"""
    f = Finding("F5", "Milestones bold scope", "NOT_APPLICABLE")

    table = find_table_after_element(tree, "MILESTONES DETAILS")
    if table is None:
        table = find_table_after_element(tree, "MILESTONE")
    if table is None:
        f.details.append("No milestones table found")
        return f

    rows = table.findall(".//tr")
    if not rows:
        f.details.append("Table has no rows")
        return f

    f.status = "PASS"
    for i, row in enumerate(rows):
        cells = row.findall(".//td")
        if not cells:
            continue
        # The milestone description is typically in the 2nd cell (index 1)
        for cell in cells:
            cell_text = (cell.text_content() or "").strip()
            if not cell_text or len(cell_text) < 10:
                continue
            bolds = get_bold_items_in_element(cell)
            if not bolds:
                continue
            total_bold = " ".join(b["bold_text"] for b in bolds)
            has_after = any(b["after_text"] for b in bolds)
            if len(total_bold) > 0 and len(total_bold) >= len(cell_text) * 0.9:
                f.status = "FAIL"
                f.details.append(f"  Row {i}: entire cell bold: \"{total_bold[:80]}\"")
            elif has_after:
                f.details.append(f"  Row {i}: partial bold OK: \"{total_bold[:50]}\" + plain: \"{bolds[0]['after_text'][:30]}\"")
            else:
                # Bold covers all text but there may be tail-less structure
                f.details.append(f"  Row {i}: bold: \"{total_bold[:60]}\" ({len(total_bold)}/{len(cell_text)} chars)")
    if not f.details:
        f.details.append("No milestone description cells found with bold")
        f.status = "INCONCLUSIVE"
    return f


def check_f8_appendix_e_heading(tree: HtmlElement, source_md: str | None) -> Finding:
    """F8: What heading level does pandoc assign to 'Appendix E'?"""
    f = Finding("F8", "Appendix E heading level", "NOT_APPLICABLE")

    for el in tree.iter():
        lvl = _heading_level(el.tag)
        if lvl is not None and "appendix e" in (el.text_content() or "").lower():
            f.details.append(f"  Pandoc rendered as <{el.tag}>: \"{el.text_content().strip()[:80]}\"")
            if lvl == 1:
                f.status = "INCONCLUSIVE"
                f.details.append("  Both pandoc and mammoth show h1 — verify in Word if numbering makes it appear as sub-heading")
            else:
                f.status = "PASS"
                f.details.append(f"  Heading level {lvl} looks correct")
            return f

    # Check if it appears as a list item instead
    for el in tree.iter("li"):
        if "appendix e" in (el.text_content() or "").lower():
            f.status = "PASS"
            f.details.append(f"  Rendered as list item: \"{el.text_content().strip()[:80]}\"")
            return f

    f.details.append("  'Appendix E' not found in document")
    return f


def check_f10_exhibits_bold(tree: HtmlElement, source_md: str | None) -> Finding:
    """F10: Exhibits table bold scope — are Exhibit D/E descriptions partially or fully bold?"""
    f = Finding("F10", "Exhibits table bold scope", "NOT_APPLICABLE")

    # Search all tables for rows containing Exhibit D/E
    for table in tree.iter("table"):
        for row in table.findall(".//tr"):
            cells = row.findall(".//td")
            # Find the row by checking if any cell contains "Exhibit D" or "Exhibit E"
            row_has_exhibit = False
            for cell in cells:
                text = (cell.text_content() or "").strip().lower()
                if text in ("exhibit d", "exhibit e"):
                    row_has_exhibit = True
                    break
            if not row_has_exhibit:
                continue

            # Check the DESCRIPTION cell (typically the next cell after the label)
            for cell in cells:
                cell_text = (cell.text_content() or "").strip()
                # Skip the label cell itself and empty cells
                if cell_text.lower() in ("exhibit d", "exhibit e", "yes", "no", ""):
                    continue
                if len(cell_text) < 10:
                    continue
                bolds = get_bold_items_in_element(cell)
                if not bolds:
                    f.details.append(f"  Desc \"{cell_text[:60]}\" — no bold")
                    continue
                total_bold = " ".join(b["bold_text"] for b in bolds)
                has_plain_tail = any(b["after_text"] for b in bolds)

                if len(total_bold) >= len(cell_text) * 0.8 and not has_plain_tail:
                    f.status = "FAIL"
                    f.details.append(f"  FULL bold: \"{total_bold[:80]}\"")
                elif has_plain_tail:
                    if f.status != "FAIL":
                        f.status = "PASS"
                    f.details.append(f"  Partial bold OK: bold=\"{bolds[0]['bold_text'][:40]}\" + plain=\"{bolds[0]['after_text'][:30]}\"")
                else:
                    if f.status != "FAIL":
                        f.status = "PASS"
                    f.details.append(f"  Bold=\"{total_bold[:50]}\" in cell=\"{cell_text[:50]}\"")

    if not f.details:
        f.details.append("No Exhibit D/E description cells found")
    return f


def check_f14_bold_colon(tree: HtmlElement, source_md: str | None) -> Finding:
    """F14: Bold boundary on colons — is the colon inside or outside <strong>?"""
    f = Finding("F14", "Bold colon boundary", "NOT_APPLICABLE")

    # Scan all bold elements in the In-Flight section for "label:" pattern
    in_section = False
    inside_count = 0
    outside_count = 0
    total_checked = 0
    mismatches: list[str] = []

    for el in tree.iter():
        lvl = _heading_level(el.tag)
        if lvl is not None and "in-flight" in (el.text_content() or "").lower():
            in_section = True
            continue
        if in_section and lvl is not None and lvl <= 2:
            break
        if not in_section:
            continue

        if el.tag == "strong":
            bold_text = (el.text_content() or "").strip()
            tail = (el.tail or "").lstrip()
            if len(bold_text) < 5:
                continue

            colon_inside = bold_text.endswith(":")
            colon_outside = tail.startswith(":")

            if colon_inside or colon_outside:
                total_checked += 1
                pos = "inside" if colon_inside else "outside"
                if colon_inside:
                    inside_count += 1
                else:
                    outside_count += 1

                # Check against source MD if available
                if source_md and colon_outside:
                    # Look for this label in source MD to see if colon should be inside
                    label_esc = re.escape(bold_text.rstrip(":").strip())
                    md_pat = re.search(rf"\*\*{label_esc}:\*\*", source_md, re.IGNORECASE)
                    if md_pat:
                        mismatches.append(f"  \"{bold_text}\" — colon outside bold, source has it inside")

    if total_checked == 0:
        f.details.append("No bold labels with colons found in In-Flight section")
        return f

    f.details.append(f"  {total_checked} bold labels checked: {inside_count} colon inside, {outside_count} colon outside")
    if mismatches:
        f.status = "FAIL"
        f.details.append(f"  {len(mismatches)} mismatch(es) vs source MD:")
        f.details.extend(mismatches[:10])
    else:
        f.status = "PASS"
        f.details.append("  All colon positions match source MD (or no source to compare)")
    return f


def check_fnew4_smart_quotes(tree: HtmlElement, source_md: str | None) -> Finding:
    """F-new4: Smart quotes — what quote characters appear?"""
    f = Finding("F-new4", "Smart quotes", "NOT_APPLICABLE")

    html_text = tree.text_content() or ""
    curly_singles = [m.start() for m in re.finditer("[\u2018\u2019]", html_text)]
    curly_doubles = [m.start() for m in re.finditer("[\u201c\u201d]", html_text)]

    f.details.append(f"  Curly single quotes found: {len(curly_singles)}")
    f.details.append(f"  Curly double quotes found: {len(curly_doubles)}")

    if source_md:
        md_curly_s = len(re.findall("[\u2018\u2019]", source_md))
        md_curly_d = len(re.findall("[\u201c\u201d]", source_md))
        f.details.append(f"  Source MD curly singles: {md_curly_s}, doubles: {md_curly_d}")

        if len(curly_singles) == 0 and len(curly_doubles) == 0:
            f.status = "PASS"
            f.details.append("  No curly quotes in HTML — no issue")
        elif md_curly_s > 0 or md_curly_d > 0:
            f.status = "PASS"
            f.details.append("  Source MD also has curly quotes — pass-through, not a defect")
        else:
            f.status = "FAIL"
            f.details.append("  HTML has curly quotes but source MD does not — potential defect")
    else:
        f.status = "INCONCLUSIVE"
        f.details.append("  No source MD to compare against")

    return f


def check_a9_bold_twelve(tree: HtmlElement, source_md: str | None) -> Finding:
    """A9: Is 'twelve (12) weeks' bold?"""
    f = Finding("A9", 'Bold on "twelve (12) weeks"', "NOT_APPLICABLE")

    html_text = tree.text_content() or ""
    if "twelve (12) weeks" not in html_text.lower() and "twelve(12)" not in html_text.lower():
        # Try variations
        if "twelve" not in html_text.lower():
            f.details.append("  Text 'twelve (12) weeks' not found in document")
            return f

    # Search for it inside <strong>
    for strong in tree.iter("strong"):
        text = (strong.text_content() or "").lower()
        if "twelve" in text and "12" in text:
            f.status = "FAIL"
            f.details.append(f"  Found in <strong>: \"{strong.text_content().strip()[:80]}\"")
            return f

    # Check if it exists as plain text
    if "twelve" in html_text.lower():
        f.status = "PASS"
        f.details.append("  'twelve (12) weeks' found as plain text (not bold) — round-trip artifact")
    else:
        f.details.append("  Text not found")
    return f


def check_a10_double_spaces(html_raw: str, source_md: str | None) -> Finding:
    """A10: Are double spaces preserved in the HTML output?"""
    f = Finding("A10", "Double spaces", "NOT_APPLICABLE")

    if source_md is None:
        f.details.append("No source MD to check for double spaces")
        f.status = "INCONCLUSIVE"
        return f

    # Find double-space locations in source MD prose (exclude tables, code, images, blank lines)
    md_lines_with_double = []
    in_code = False
    for i, line in enumerate(source_md.splitlines(), 1):
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        stripped = line.strip()
        # Skip MD table rows, image refs, blank lines, horizontal rules, and
        # lines where "  " is just MD indentation (leading spaces before list markers)
        if stripped.startswith("|") or stripped.startswith("![") or not stripped or stripped.startswith("---"):
            continue
        # Check for double spaces in the CONTENT (not leading whitespace)
        if "  " not in stripped:
            continue
        idx = stripped.index("  ")
        context = stripped[max(0, idx - 20):idx + 22]
        md_lines_with_double.append((i, context))

    if not md_lines_with_double:
        f.status = "PASS"
        f.details.append("  No double spaces found in source MD")
        return f

    f.details.append(f"  Source MD has {len(md_lines_with_double)} lines with double spaces")

    # Check if HTML preserves them (raw HTML, not rendered)
    # Look for double spaces or &nbsp; sequences
    html_double_spaces = len(re.findall(r"(?<!\s)  (?!\s)", html_raw))
    html_nbsp = html_raw.count("&nbsp;")

    f.details.append(f"  HTML raw double spaces: {html_double_spaces}")
    f.details.append(f"  HTML &nbsp; count: {html_nbsp}")

    if html_double_spaces > 0 or html_nbsp > 0:
        f.status = "PASS"
        f.details.append("  Double spaces preserved in HTML — round-trip artifact (MarkItDown collapsed them)")
    else:
        f.status = "FAIL"
        f.details.append("  Double spaces NOT preserved — may be a real defect in DOCX")

    # Show a few examples from source
    for line_no, ctx in md_lines_with_double[:3]:
        f.details.append(f"  Example (MD line {line_no}): \"{ctx}\"")

    return f


# ---------------------------------------------------------------------------
# Check registry
# ---------------------------------------------------------------------------

CHECKS: dict[str, tuple] = {
    "F5":     (check_f5_milestones_bold,    ["Fiserv"]),
    "F8":     (check_f8_appendix_e_heading, ["Fiserv"]),
    "F10":    (check_f10_exhibits_bold,     ["Fiserv"]),
    "F14":    (check_f14_bold_colon,        ["Fiserv"]),
    "F-new4": (check_fnew4_smart_quotes,    ["Fiserv"]),
    "A9":     (check_a9_bold_twelve,        ["ALAS"]),
    "A10":    (check_a10_double_spaces,     ["ALAS"]),
}


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def print_report(
    docx_path: Path,
    html_path: Path,
    source_path: Path | None,
    results: list[Finding],
) -> None:
    print("\n=== Pandoc HTML Validation Report ===")
    print(f"Document: {docx_path.name}")
    print(f"HTML saved: {html_path.name}")
    if source_path:
        print(f"Source MD: {source_path.name}")
    print()
    print("--- Issue Checks ---")

    real_defects = []
    artifacts = []
    inconclusive = []
    not_applicable = []

    for r in results:
        print(f"\n{r.issue_id} ({r.title}): {r.status}")
        for d in r.details:
            print(d)

        if r.status == "FAIL":
            real_defects.append(r.issue_id)
        elif r.status == "PASS":
            artifacts.append(r.issue_id)
        elif r.status == "INCONCLUSIVE":
            inconclusive.append(r.issue_id)
        else:
            not_applicable.append(r.issue_id)

    print("\n--- Summary ---")
    print(f"Likely real defects:    {', '.join(real_defects) or 'none'}")
    print(f"Round-trip artifacts:   {', '.join(artifacts) or 'none'}")
    print(f"Inconclusive:          {', '.join(inconclusive) or 'none'}")
    print(f"Not applicable:        {', '.join(not_applicable) or 'none'}")
    print()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Pandoc HTML validation for md2word output DOCX files",
    )
    parser.add_argument("docx", type=Path, help="Path to the output DOCX file")
    parser.add_argument("--source", type=Path, default=None, help="Path to source MD for comparison")
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"Error: {args.docx} not found", file=sys.stderr)
        sys.exit(1)

    source_md: str | None = None
    if args.source:
        if not args.source.exists():
            print(f"Error: {args.source} not found", file=sys.stderr)
            sys.exit(1)
        source_md = args.source.read_text(encoding="utf-8")

    # Determine which document set this is
    doc_name = args.docx.stem.lower()
    if "fiserv" in doc_name:
        doc_set = "Fiserv"
    elif "alas" in doc_name:
        doc_set = "ALAS"
    else:
        doc_set = "unknown"

    # Convert
    print(f"Converting {args.docx.name} to HTML via pandoc...")
    html_content = convert_docx_to_html(args.docx)
    html_path = save_html(html_content, args.docx)
    print(f"HTML saved to {html_path.name}")

    # Parse
    tree = lhtml.fromstring(html_content)

    # Run applicable checks
    results: list[Finding] = []
    for issue_id, (check_fn, applicable_docs) in CHECKS.items():
        if doc_set not in applicable_docs and "all" not in applicable_docs:
            # Mark as not applicable but still include
            results.append(Finding(issue_id, check_fn.__doc__.split(":")[0] if check_fn.__doc__ else issue_id, "NOT_APPLICABLE", ["  Not applicable to this document"]))
            continue

        if check_fn == check_a10_double_spaces:
            result = check_fn(html_content, source_md)
        else:
            result = check_fn(tree, source_md)
        results.append(result)

    print_report(args.docx, html_path, args.source, results)


if __name__ == "__main__":
    main()
