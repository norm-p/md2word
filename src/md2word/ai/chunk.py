"""
chunk.py — Split DOCX OOXML into logical sections.

Parses word/document.xml and groups consecutive elements under each heading,
so downstream AI steps (map, edit_plan) work on manageable fragments rather
than the full XML blob.

Heading detection is style-aware: reads styles.xml to build a map of all
paragraph styles that function as structural headings, including custom styles
that inherit from built-in Heading or outline-level styles.

Tables (w:tbl), structured document tags (w:sdt), and other non-paragraph
elements are included in their enclosing section's fragment so the AI sees
the full section content.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

_W = NS["w"]
_HEADING_ID_RE = re.compile(r"^[Hh]eading\s*(\d+)$")
# Style names that suggest a heading role (matched against style ID and display name)
_HEADING_NAME_RE = re.compile(r"\b(?:heading|title|article)\b", re.IGNORECASE)
# List style patterns — match against style ID and display name
_LIST_BULLET_RE = re.compile(r"^(?:[Ll]ist[\-\s]?[Bb]ullet|[Bb]ullets?$)", re.IGNORECASE)
_LIST_NUMBER_RE = re.compile(r"^[Ll]ist[\-\s]?[Nn]umber", re.IGNORECASE)

# ---------------------------------------------------------------------------
# Boilerplate separator detection (dynamic)
# ---------------------------------------------------------------------------
# A paragraph is treated as a boilerplate separator when it looks like a
# document-level divider: short, all-caps, bold, non-heading.  This catches
# "SIGNATURE PAGE FOLLOWS", "END OF DOCUMENT", etc. without hardcoding.

_SEPARATOR_MIN_LEN = 8
_SEPARATOR_MAX_LEN = 80


def _is_boilerplate_separator(para: etree._Element, heading_map: dict[str, int] | None) -> bool:
    """True if a w:p element looks like a boilerplate section divider.

    Criteria (all must be met):
    - Not a heading-styled paragraph
    - Text is 8-80 characters of substantially all-caps alphabetic content
    - Paragraph has bold formatting (w:b in run properties)
    """
    # Must not be a heading
    if _para_heading_level(para, heading_map) is not None:
        return False

    text = _para_text(para).strip()
    if not (_SEPARATOR_MIN_LEN <= len(text) <= _SEPARATOR_MAX_LEN):
        return False

    # Must be substantially all-uppercase (allow digits, punctuation, spaces)
    alpha = [c for c in text if c.isalpha()]
    if len(alpha) < 4 or not all(c.isupper() for c in alpha):
        return False

    # Must have bold formatting
    if para.find(f".//{{{_W}}}b") is None:
        return False

    return True


@dataclass
class DocxSection:
    heading: str          # text of the heading paragraph
    heading_level: int    # 1-9 (0 for preamble)
    xml_fragment: str     # raw XML for all elements in this section (heading included)


def extract_document_xml(docx_path: Path) -> str:
    """Read word/document.xml from a .docx archive."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return zf.read("word/document.xml").decode("utf-8")


def extract_styles_xml(docx_path: Path) -> str | None:
    """Read word/styles.xml from a .docx archive, or None if absent."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        if "word/styles.xml" not in zf.namelist():
            return None
        return zf.read("word/styles.xml").decode("utf-8")


def build_heading_style_map(styles_xml: str) -> dict[str, int]:
    """Return {styleId: headingLevel} for all paragraph styles that act as headings.

    headingLevel is 1-9. A style qualifies as a heading if:
    1. Its ID matches the built-in HeadingN pattern (e.g. "Heading1"), or
    2. It has a direct w:outlineLvl of 0-8 (outline-level styles like ARTICLEBL*), or
    3. Its display name or ID contains heading-related words AND it inherits from
       a qualifying ancestor (covers custom styles like "ArticleHeading", "Title").

    Rule 3 requires the name check to avoid promoting body-text styles (e.g.,
    "Body-Numbered") that happen to inherit from an outline-level base.
    """
    root = etree.fromstring(styles_xml.encode("utf-8"))

    style_info: dict[str, dict] = {}
    for style in root.findall(f".//{{{_W}}}style"):
        if style.get(f"{{{_W}}}type") != "paragraph":
            continue
        sid = style.get(f"{{{_W}}}styleId", "")
        if not sid:
            continue

        name_el = style.find(f"{{{_W}}}name")
        name = name_el.get(f"{{{_W}}}val", "") if name_el is not None else ""

        lvl_el = style.find(f".//{{{_W}}}outlineLvl")
        outline_lvl = int(lvl_el.get(f"{{{_W}}}val", "99")) if lvl_el is not None else 99

        based_el = style.find(f"{{{_W}}}basedOn")
        based_on = based_el.get(f"{{{_W}}}val", "") if based_el is not None else ""

        style_info[sid] = {"name": name, "outlineLvl": outline_lvl, "basedOn": based_on}

    result: dict[str, int] = {}

    def resolve(sid: str, visited: frozenset = frozenset()) -> int | None:
        if sid in result:
            return result[sid]
        if sid in visited or sid not in style_info:
            return None

        info = style_info[sid]

        # Rule 1: built-in HeadingN style ID
        m = _HEADING_ID_RE.match(sid)
        if m:
            result[sid] = int(m.group(1))
            return result[sid]

        # Rule 2: has direct w:outlineLvl 0-8 — definitively a structural heading
        if info["outlineLvl"] <= 8:
            result[sid] = info["outlineLvl"] + 1
            return result[sid]

        # Rule 3: name/ID hints "heading"/"title"/"article" and inherits from a heading
        if _HEADING_NAME_RE.search(info["name"]) or _HEADING_NAME_RE.search(sid):
            parent = info["basedOn"]
            if parent:
                parent_level = resolve(parent, visited | {sid})
                if parent_level is not None:
                    result[sid] = parent_level
                    return parent_level

        # Rule 4: directly inherits from a known heading style (no name check required).
        # Catches document-specific styles like "Body-Numbered" (basedOn ARTICLEBL2)
        # that function as sub-section headings but use generic names.
        parent = info["basedOn"]
        if parent:
            parent_level = resolve(parent, visited | {sid})
            if parent_level is not None:
                result[sid] = parent_level
                return parent_level

        return None

    for sid in list(style_info):
        resolve(sid)

    return result


def build_list_style_map(styles_xml: str) -> dict[str, str | None]:
    """Return {"bullet": style_id, "numbered": style_id} for list paragraph styles.

    Searches styles.xml for paragraph styles whose ID or display name matches
    canonical list-bullet or list-number naming patterns. Returns None for a
    key if no matching style is found — callers should warn the user and fall
    back to "ListBullet" / "ListNumber" as Word built-in defaults.

    When multiple candidates exist, the one with the shortest ID is preferred
    (e.g. "ListBullet" over "ListBullet2") as it is typically the base style.
    """
    root = etree.fromstring(styles_xml.encode("utf-8"))

    bullet_styles: list[str] = []
    numbered_styles: list[str] = []

    for style in root.findall(f".//{{{_W}}}style"):
        if style.get(f"{{{_W}}}type") != "paragraph":
            continue
        sid = style.get(f"{{{_W}}}styleId", "")
        if not sid:
            continue
        name_el = style.find(f"{{{_W}}}name")
        name = name_el.get(f"{{{_W}}}val", "") if name_el is not None else ""

        if _LIST_BULLET_RE.match(sid) or _LIST_BULLET_RE.match(name):
            bullet_styles.append(sid)
        elif _LIST_NUMBER_RE.match(sid) or _LIST_NUMBER_RE.match(name):
            numbered_styles.append(sid)

    # Prefer the base style (shortest ID = least-specific, e.g. ListBullet over ListBullet2)
    bullet = min(bullet_styles, key=len) if bullet_styles else None
    numbered = min(numbered_styles, key=len) if numbered_styles else None

    return {"bullet": bullet, "numbered": numbered}


def _para_text(para: etree._Element) -> str:
    """Extract plain text from a w:p element."""
    return "".join(t.text or "" for t in para.findall(f".//{{{_W}}}t"))


def _para_heading_level(
    para: etree._Element,
    heading_map: dict[str, int] | None,
) -> int | None:
    """Return heading level (1-9) if this paragraph has a heading style, else None.

    Uses heading_map (from build_heading_style_map) when provided, falling back
    to the built-in HeadingN regex for documents without a custom style map.
    """
    ppr = para.find(f"{{{_W}}}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{{{_W}}}pStyle")
    if pstyle is None:
        return None
    val = pstyle.get(f"{{{_W}}}val", "")

    if heading_map is not None:
        return heading_map.get(val)

    # Fallback: built-in HeadingN regex only
    m = _HEADING_ID_RE.match(val)
    return int(m.group(1)) if m else None


def _serialize(element: etree._Element) -> str:
    """Serialize an XML element to string, preserving namespaces."""
    return etree.tostring(element, encoding="unicode")


def chunk_docx_xml(
    document_xml: str,
    heading_map: dict[str, int] | None = None,
) -> list[DocxSection]:
    """Parse document.xml and return a list of DocxSection objects.

    Each section starts at a heading paragraph and ends just before the next
    heading (or end of document). Tables, structured-document tags, and other
    block-level elements are included in the enclosing section's fragment.

    Paragraphs before the first heading are grouped under a synthetic
    "(preamble)" section at level 0.

    Pass heading_map (from build_heading_style_map) to recognise custom heading
    styles. Without it, only built-in Heading1-9 styles are detected.
    """
    root = etree.fromstring(document_xml.encode("utf-8"))
    body = root.find(f"{{{_W}}}body")
    if body is None:
        return []

    sections: list[DocxSection] = []
    current_heading = "(preamble)"
    current_level = 0
    current_elems: list[etree._Element] = []

    for child in body:
        # lxml marks comments/PIs with a callable tag — skip them
        if callable(child.tag):
            continue

        local = etree.QName(child.tag).localname

        # w:sectPr is document-level section properties, not content
        if local == "sectPr":
            continue

        if local == "p":
            level = _para_heading_level(child, heading_map)
            if level is not None:
                # Flush previous section
                if current_elems:
                    sections.append(DocxSection(
                        heading=current_heading,
                        heading_level=current_level,
                        xml_fragment="".join(_serialize(e) for e in current_elems),
                    ))
                current_heading = _para_text(child)
                current_level = level
                current_elems = [child]
                continue

            # Check for boilerplate separators (e.g. bold all-caps "SIGNATURE
            # PAGE FOLLOWS"). These create a synthetic section break so trailing
            # boilerplate is preserved as-is rather than regenerated by AI.
            if current_elems and _is_boilerplate_separator(child, heading_map):
                sections.append(DocxSection(
                    heading=current_heading,
                    heading_level=current_level,
                    xml_fragment="".join(_serialize(e) for e in current_elems),
                ))
                # Use a normalized synthetic heading so both parsers agree
                sep_text = _para_text(child).strip()
                current_heading = f"({sep_text.lower()})"
                current_level = 0
                current_elems = [child]
                continue

        # w:tbl, w:sdt, w:p (non-heading), and anything else go into current section
        current_elems.append(child)

    # Flush last section
    if current_elems:
        sections.append(DocxSection(
            heading=current_heading,
            heading_level=current_level,
            xml_fragment="".join(_serialize(e) for e in current_elems),
        ))

    return sections
