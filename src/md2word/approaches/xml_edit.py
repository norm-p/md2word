"""
xml_edit.py — Apply an EditPlan to an existing .docx via XML surgery (update mode).

Steps:
  1. Scan for tracked changes/comments (deterministic)
  2. Chunk document.xml into sections (deterministic)
  3. Map MD sections to DOCX sections (AI)
  4. Build edit plan (AI, batched; validates/retries XML per fragment)
  5. Validate media relationship references (deterministic)
  6. Summarize changes and confirm (AI, unless --accept-changes)
  7. Whole-document XML validation (deterministic)
  8. Repackage .docx preserving zip metadata (deterministic)
"""

from __future__ import annotations

import copy
import os
import re
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path

import click
from lxml import etree

from ..ai.chunk import NS, _is_boilerplate_separator, build_heading_style_map, build_list_style_map, chunk_docx_xml, extract_document_xml, extract_styles_xml
from ..ai.client import get_client_or_none
from ..ai.conflict import detect_conflicts
from ..ai.edit_plan import Edit, build_edit_plan
from ..ai.map import SectionMapping, map_sections, map_sections_deterministic
from ..ai.summarize import summarize_changes

_W = NS["w"]
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_BULLET_RE = re.compile(r"^[ \t]*[-*]\s+", re.MULTILINE)

# Markdown formatting patterns stripped during text normalization
_MD_STRIP_RE = re.compile(r"[*_#`\[\]()!|]")
_WHITESPACE_RE = re.compile(r"\s+")

# Patterns for Option B (additive bullet patch) and Option C (in-place text correction)
_MD_BULLET_RE = re.compile(r"^[ \t]*[-*+]\s+(.+?)[ \t]*$", re.MULTILINE)
_MD_INLINE_RE = re.compile(r"\*{1,2}|_{1,2}|`{1,3}|~~")
_MD_TABLE_CELL_RE = re.compile(r"\*{1,2}|`{1,3}|~~")  # like _MD_INLINE_RE but keeps underscores

# Module-level heading map — populated from styles.xml at run() startup.
# Used by _para_heading_text to support custom heading styles.
_heading_map: dict[str, int] | None = None


# ---------------------------------------------------------------------------
# Deterministic unchanged detection
# ---------------------------------------------------------------------------

def _normalize_text(text: str) -> str:
    """Normalize text for comparison: strip formatting, collapse whitespace, lowercase."""
    # Normalize unicode (curly quotes → straight, em-dashes → hyphens, etc.)
    text = unicodedata.normalize("NFKD", text)
    # Strip markdown formatting characters
    text = _MD_STRIP_RE.sub("", text)
    # Collapse whitespace
    text = _WHITESPACE_RE.sub(" ", text).strip().lower()
    return text


def _extract_docx_section_text(xml_fragment: str) -> str:
    """Extract plain text from a DOCX section's XML fragment."""
    try:
        wrapped = f'<root xmlns:w="{_W}">{xml_fragment}</root>'
        root = etree.fromstring(wrapped.encode("utf-8"))
        texts: list[str] = []
        for elem in root.iter():
            if callable(elem.tag):
                continue
            local = etree.QName(elem.tag).localname
            if local == "t" and elem.text:
                texts.append(elem.text)
            elif local == "p":
                texts.append("\n")
        return " ".join(texts)
    except Exception:
        return ""


def _sections_text_match(mapping: SectionMapping) -> bool:
    """Return True if the MD content and DOCX section have equivalent text.

    When text is essentially the same (after stripping formatting), there is
    no reason to regenerate the XML — the original DOCX formatting should be
    preserved as-is.
    """
    if mapping.docx_section is None or mapping.action != "replace":
        return False

    md_norm = _normalize_text(mapping.md_content)
    docx_norm = _normalize_text(_extract_docx_section_text(mapping.docx_section.xml_fragment))

    if not md_norm or not docx_norm:
        return False

    # Exact match after normalization
    if md_norm == docx_norm:
        return True

    # If MD content has a Markdown table but the DOCX section has no w:tbl,
    # the section is structurally incomplete — force replace so the LLM adds it.
    if "|" in mapping.md_content:
        try:
            wrapped = f'<root xmlns:w="{_W}">{mapping.docx_section.xml_fragment}</root>'
            frag_root = etree.fromstring(wrapped.encode("utf-8"))
            if frag_root.find(f"{{{_W}}}tbl") is None:
                return False
        except Exception:
            pass

    # Character-level similarity via longest common subsequence ratio.
    # No length pre-check: MD and DOCX representations of the same content can
    # legitimately differ in length by 10-20% due to OOXML whitespace, run
    # splitting, and list numbering — a length ratio gate causes false negatives.
    from difflib import SequenceMatcher
    ratio = SequenceMatcher(None, md_norm, docx_norm, autojunk=False).ratio()

    # Disqualify if the heading paragraph has a heading style but empty text —
    # the LLM must regenerate so R2-B heading rescue can inject the heading text.
    if ratio >= 0.90:
        try:
            wrapped = f'<root xmlns:w="{_W}">{mapping.docx_section.xml_fragment}</root>'
            frag_root = etree.fromstring(wrapped.encode("utf-8"))
            first_p = frag_root.find(f"{{{_W}}}p")
            if first_p is not None and _heading_map is not None:
                ppr = first_p.find(f"{{{_W}}}pPr")
                if ppr is not None:
                    pstyle = ppr.find(f"{{{_W}}}pStyle")
                    if pstyle is not None:
                        style_id = pstyle.get(f"{{{_W}}}val", "")
                        if style_id in _heading_map:
                            run_texts = "".join(
                                t.text or "" for t in first_p.findall(f".//{{{_W}}}t")
                            )
                            if not run_texts.strip():
                                return False
        except Exception:
            pass

    if ratio >= 0.90:
        # Fix 4: Bullet-count structural guard — if MD has many more bullets
        # than DOCX, the section has significant new content that blob similarity
        # misses. Force LLM replacement as safety net.
        md_bullet_count = len(_MD_BULLET_RE.findall(mapping.md_content))
        docx_list_count = mapping.docx_section.xml_fragment.count("<w:numPr")
        if md_bullet_count > docx_list_count + 5:
            return False
        return True
    return False


# ---------------------------------------------------------------------------
# Heading helpers
# ---------------------------------------------------------------------------

def _extract_heading_styles_in_use(docx_sections) -> dict[str, int]:
    """Return {styleId: level} for heading styles actually used in the existing document.

    Reads the first paragraph of each non-preamble section to find the style ID
    that the document author actually applied, then picks the most frequently used
    style per level. This is more reliable than the full heading_map, which can
    include incidental styles (e.g. Title, TOC Heading) at the same outline level
    as the primary heading style (e.g. ArticleHeading).
    """
    from collections import Counter
    counts: Counter = Counter()       # style_id → occurrence count
    levels: dict[str, int] = {}      # style_id → heading level
    for section in docx_sections:
        if section.heading_level == 0:
            continue
        try:
            wrapped = f'<root xmlns:w="{_W}">{section.xml_fragment}</root>'
            root = etree.fromstring(wrapped.encode("utf-8"))
            first_p = root.find(f"{{{_W}}}p")
            if first_p is not None:
                ppr = first_p.find(f"{{{_W}}}pPr")
                if ppr is not None:
                    pstyle = ppr.find(f"{{{_W}}}pStyle")
                    if pstyle is not None:
                        style_id = pstyle.get(f"{{{_W}}}val", "")
                        if style_id:
                            counts[style_id] += 1
                            levels[style_id] = section.heading_level
        except Exception:
            pass

    # For each level, pick the most frequently used style
    best: dict[int, tuple[str, int]] = {}  # level → (style_id, count)
    for style_id, count in counts.items():
        lvl = levels[style_id]
        if lvl not in best or count > best[lvl][1]:
            best[lvl] = (style_id, count)

    return {style_id: lvl for lvl, (style_id, _) in best.items()}

def _para_heading_text(para: etree._Element) -> str | None:
    """Return heading text if this paragraph is a heading or boilerplate separator.

    Returns heading text for heading-styled paragraphs, or a normalized
    synthetic heading like ``(signature page follows)`` for boilerplate
    separator paragraphs (matching chunk.py's dynamic detection).
    """
    # Check boilerplate separator first (these have no heading style)
    if _is_boilerplate_separator(para, _heading_map):
        text = "".join(t.text or "" for t in para.findall(f".//{{{_W}}}t"))
        return f"({text.strip().lower()})"

    ppr = para.find(f"{{{_W}}}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{{{_W}}}pStyle")
    if pstyle is None:
        return None
    val = pstyle.get(f"{{{_W}}}val", "")

    if _heading_map is not None:
        if val not in _heading_map:
            return None
    else:
        import re
        if not re.match(r"^[Hh]eading\s*(\d+)$", val):
            return None

    return "".join(t.text or "" for t in para.findall(f".//{{{_W}}}t"))


# ---------------------------------------------------------------------------
# XML fragment helpers
# ---------------------------------------------------------------------------

def _parse_xml_fragment(xml_str: str) -> list[etree._Element]:
    """Parse an XML string of w:p / w:tbl elements into a list of Elements."""
    wrapped = f'<root xmlns:w="{_W}">{xml_str}</root>'
    root = etree.fromstring(wrapped.encode("utf-8"))
    return list(root)


def _validate_document_xml(xml_bytes: bytes) -> None:
    """Confirm the assembled document.xml is well-formed. Raises on corruption."""
    try:
        etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"Assembled document.xml is not well-formed XML: {exc}") from exc


# ---------------------------------------------------------------------------
# Section range lookup
# ---------------------------------------------------------------------------

def _find_section_range(body: etree._Element, heading_text: str) -> tuple[int, int]:
    """Find the [start, end) indices in body's child list for a named section.

    Includes all element types (w:p, w:tbl, w:sdt, …) between the heading
    paragraph and the next heading paragraph. Stops before w:sectPr.
    """
    # "(preamble)" is a synthetic name for content before the first heading.
    # Map it to the range from body index 0 to the first heading paragraph.
    if heading_text == "(preamble)":
        children = list(body)
        end = len(children)
        for i, child in enumerate(children):
            if callable(child.tag):
                continue
            if child.tag == f"{{{_W}}}sectPr":
                end = i
                break
            if child.tag == f"{{{_W}}}p" and _para_heading_text(child) is not None:
                end = i
                break
        return 0, end

    children = list(body)
    start = None

    for i, child in enumerate(children):
        if callable(child.tag):
            continue
        if child.tag == f"{{{_W}}}p":
            ht = _para_heading_text(child)
            if ht == heading_text:
                start = i
                break

    if start is None:
        raise ValueError(f"Section not found: {heading_text!r}")

    end = len(children)
    for i in range(start + 1, len(children)):
        if callable(children[i].tag):
            continue
        tag = children[i].tag
        # Stop at document section properties
        if tag == f"{{{_W}}}sectPr":
            end = i
            break
        # Stop at next heading paragraph
        if tag == f"{{{_W}}}p" and _para_heading_text(children[i]) is not None:
            end = i
            break

    return start, end


# ---------------------------------------------------------------------------
# Tracked changes acceptance (O(n) with lxml .getparent())
# ---------------------------------------------------------------------------

def _accept_tracked_changes(document_xml: str) -> str:
    """Accept all tracked changes: merge w:ins content and remove w:del."""
    root = etree.fromstring(document_xml.encode("utf-8"))

    # Accept insertions: lift children out of w:ins in place
    for ins in root.findall(f".//{{{_W}}}ins"):
        parent = ins.getparent()
        if parent is None:
            continue
        idx = list(parent).index(ins)
        children = list(ins)
        parent.remove(ins)
        for j, child in enumerate(children):
            parent.insert(idx + j, child)

    # Remove deletions entirely
    for del_elem in root.findall(f".//{{{_W}}}del"):
        parent = del_elem.getparent()
        if parent is not None:
            parent.remove(del_elem)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")


# ---------------------------------------------------------------------------
# Apply edits
# ---------------------------------------------------------------------------

def _apply_edits(document_xml: str, edits: list[Edit]) -> bytes:
    """Apply a list of Edits to document.xml and return the modified XML bytes.

    Edits are processed in reverse order. This keeps replace/delete indices stable
    as earlier sections are modified, and also ensures that multiple inserts
    targeting the same anchor appear in the correct MD source order: the last
    insert (processed first in reverse) lands immediately after the anchor,
    and each earlier insert then slots in before it, building up the correct sequence.
    """
    root = etree.fromstring(document_xml.encode("utf-8"))
    body = root.find(f"{{{_W}}}body")
    if body is None:
        raise ValueError("No w:body found in document.xml")

    deferred_inserts: list[tuple[Edit, list[etree._Element]]] = []

    for edit in reversed(edits):
        new_elements = _parse_xml_fragment(edit.content) if edit.content else []

        if edit.kind == "delete":
            start, end = _find_section_range(body, edit.target_heading)
            children = list(body)
            for child in children[start:end]:
                body.remove(child)

        elif edit.kind == "replace":
            start, end = _find_section_range(body, edit.target_heading)
            children = list(body)
            # Collect drawing paragraphs from original section before removal so
            # they can be re-appended if the new XML contains no drawings.
            old_drawing_paras = [
                c for c in children[start:end]
                if c.tag == f"{{{_W}}}p" and c.find(f".//{{{_W}}}drawing") is not None
            ]
            # Collect tables from original section — LLM cannot reliably regenerate
            # complex OOXML table structures (merged cells, borders, column widths).
            old_tables = [
                c for c in children[start:end]
                if c.tag == f"{{{_W}}}tbl"
            ]
            # Fix A: Save original heading paragraph (carries correct style).
            # The LLM is unreliable at choosing between multiple styles at the
            # same heading level (e.g. Style1 vs TitleLevel2 at level 2).
            original_heading_para = None
            for child in children[start:end]:
                if child.tag == f"{{{_W}}}p" and _para_heading_text(child) is not None:
                    original_heading_para = copy.deepcopy(child)
                    break
            for child in children[start:end]:
                body.remove(child)
            for j, elem in enumerate(new_elements):
                body.insert(start + j, elem)
            tail_pos = start + len(new_elements)
            # Fix A: Replace LLM's heading paragraph with original (preserves
            # correct style and heading text for replace edits).
            if original_heading_para is not None:
                for k, elem in enumerate(new_elements):
                    if elem.tag == f"{{{_W}}}p" and _para_heading_text(elem) is not None:
                        body.remove(elem)
                        body.insert(start + k, original_heading_para)
                        new_elements[k] = original_heading_para
                        break
                else:
                    # LLM didn't generate a heading — prepend original
                    body.insert(start, original_heading_para)
                    new_elements.insert(0, original_heading_para)
                    tail_pos += 1
            # Fix B: Strip heading styles from non-first paragraphs. Each
            # replacement section should have exactly one heading (the section
            # heading preserved by Fix A). Additional heading-styled paragraphs
            # are false promotions by the LLM (e.g. bold text → Style2).
            heading_seen = False
            for elem in new_elements:
                if elem.tag != f"{{{_W}}}p":
                    continue
                if _para_heading_text(elem) is None:
                    continue
                if not heading_seen:
                    heading_seen = True
                    continue
                ppr = elem.find(f"{{{_W}}}pPr")
                if ppr is not None:
                    ps = ppr.find(f"{{{_W}}}pStyle")
                    if ps is not None:
                        ppr.remove(ps)
            # Fix B2: Remove LLM body paragraphs whose text matches the section
            # heading.  The LLM sometimes emits the heading as a plain paragraph
            # (without heading style) which Fix B cannot catch.  Only check the
            # first 2 body paragraphs — duplicate headings from the LLM always
            # appear near the top.  Checking deeper risks removing legitimate
            # content that happens to match the heading text.
            if original_heading_para is not None:
                orig_heading_text = _normalize_text(
                    "".join(
                        (t.text or "")
                        for t in original_heading_para.findall(f".//{{{_W}}}t")
                    )
                )
                if orig_heading_text:
                    to_remove = []
                    first = True
                    body_para_count = 0
                    for elem in new_elements:
                        if elem.tag != f"{{{_W}}}p":
                            continue
                        if first:
                            first = False
                            continue  # skip the heading itself
                        body_para_count += 1
                        if body_para_count > 2:
                            break  # only check first 2 body paragraphs
                        para_text = _normalize_text(
                            "".join(
                                (t.text or "")
                                for t in elem.findall(f".//{{{_W}}}t")
                            )
                        )
                        if para_text and para_text == orig_heading_text:
                            to_remove.append(elem)
                    for elem in to_remove:
                        body.remove(elem)
                        new_elements.remove(elem)
                        tail_pos -= 1
            # Fix B3: Remove consecutive duplicate paragraphs in LLM output.
            # The LLM sometimes emits the same sub-heading or intro paragraph
            # twice in a row (e.g. "Delivery Experience Framework" x2).
            # Compare adjacent paragraph texts; remove the second if identical.
            prev_text: str | None = None
            b3_remove: list[etree._Element] = []
            for elem in new_elements:
                if elem.tag != f"{{{_W}}}p":
                    prev_text = None
                    continue
                para_text = _normalize_text(
                    "".join(
                        (t.text or "")
                        for t in elem.findall(f".//{{{_W}}}t")
                    )
                )
                if not para_text:
                    prev_text = None
                    continue
                if para_text == prev_text:
                    b3_remove.append(elem)
                else:
                    prev_text = para_text
            for elem in b3_remove:
                body.remove(elem)
                new_elements.remove(elem)
                tail_pos -= 1
                click.echo(
                    f"  Removed consecutive duplicate paragraph in "
                    f"'{edit.target_heading}'"
                )
            # Preserve images: if original had drawings and new content has none, re-append them
            if old_drawing_paras and not any(
                e.find(f".//{{{_W}}}drawing") is not None for e in new_elements
            ):
                for j, elem in enumerate(old_drawing_paras):
                    body.insert(tail_pos + j, elem)
                tail_pos += len(old_drawing_paras)
            # Preserve tables: rescue originals when LLM output has no tables
            # or when LLM-generated tables have dramatically fewer rows (< 50%)
            # than the originals (hallucinated skeleton table).
            new_tables = [e for e in new_elements if e.tag == f"{{{_W}}}tbl"]
            def _tbl_row_count(tbl: etree._Element) -> int:
                return len(tbl.findall(f".//{{{_W}}}tr"))
            # Remove spurious LLM-generated tables when the original DOCX
            # section had no tables — the LLM hallucinated them from the
            # structural summary or surrounding context.
            if not old_tables and new_tables:
                click.echo(
                    f"  Removing {len(new_tables)} spurious LLM-generated table(s) in "
                    f"'{edit.target_heading}' — original section had none."
                )
                for bad_tbl in new_tables:
                    body.remove(bad_tbl)
                    tail_pos -= 1
                new_tables = []
            rescue_tables = False
            if old_tables and not new_tables:
                rescue_tables = True
            elif old_tables and new_tables and all(
                _tbl_row_count(nt) < _tbl_row_count(ot) * 0.5
                for nt, ot in zip(new_tables, old_tables)
            ):
                rescue_tables = True
            if rescue_tables:
                click.echo(
                    f"  Preserving {len(old_tables)} original table(s) in "
                    f"'{edit.target_heading}' — LLM output had none/insufficient rows."
                )
                for bad_tbl in new_tables:
                    body.remove(bad_tbl)
                    tail_pos -= 1
                for j, tbl in enumerate(old_tables):
                    body.insert(tail_pos + j, tbl)

        elif edit.kind == "insert":
            if edit.target_heading == "(end)":
                sect_pr = body.find(f"{{{_W}}}sectPr")
                if sect_pr is not None:
                    idx = list(body).index(sect_pr)
                    for j, elem in enumerate(new_elements):
                        body.insert(idx + j, elem)
                else:
                    for elem in new_elements:
                        body.append(elem)
            else:
                try:
                    _, end = _find_section_range(body, edit.target_heading)
                    for j, elem in enumerate(new_elements):
                        body.insert(end + j, elem)
                except ValueError:
                    # Anchor not found — defer for retry after all edits are applied.
                    # The anchor may be created by an insert processed later in this
                    # reverse pass.
                    deferred_inserts.append((edit, new_elements))

    # Retry deferred inserts in forward order (anchors may now exist).
    for edit, new_elements in deferred_inserts:
        try:
            _, end = _find_section_range(body, edit.target_heading)
            for j, elem in enumerate(new_elements):
                body.insert(end + j, elem)
            click.echo(
                f"  Deferred insert for '{edit.target_heading}' succeeded on retry."
            )
        except ValueError:
            click.echo(
                f"  Warning: insert anchor '{edit.target_heading}' not found "
                "— SKIPPING insert (content not placed)."
            )

    # Final cleanup: remove heading-styled paragraphs with no visible text.
    # These are structural artifacts (e.g., empty Heading2 at the end of a
    # section) that the section-level filter can't catch because they're
    # embedded within non-empty sections rather than being standalone sections.
    # _para_heading_text() returns "" (not None) for headings with no text.
    empty_heading_count = 0
    for child in list(body):
        if child.tag != f"{{{_W}}}p":
            continue
        heading_text = _para_heading_text(child)
        if heading_text is None:
            continue  # not a heading paragraph
        if heading_text.strip():
            continue  # heading has visible text — keep
        # It's a heading-styled paragraph with no visible text — remove it
        body.remove(child)
        empty_heading_count += 1
    if empty_heading_count:
        click.echo(
            f"  Removed {empty_heading_count} empty heading paragraph(s)"
        )

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# Media / relationship validation
# ---------------------------------------------------------------------------

def _load_relationship_ids(docx_path: Path) -> set[str]:
    """Return the set of relationship IDs declared in word/_rels/document.xml.rels."""
    rels_file = "word/_rels/document.xml.rels"
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            if rels_file not in zf.namelist():
                return set()
            rels_xml = zf.read(rels_file)
        root = etree.fromstring(rels_xml)
        return {rel.get("Id") for rel in root if rel.get("Id") is not None}
    except Exception:
        return set()


def _check_media_refs(edits: list[Edit], valid_ids: set[str]) -> None:
    """Warn about any r:id / r:embed / r:link values not found in valid_ids."""
    if not valid_ids:
        return
    for edit in edits:
        if not edit.content:
            continue
        try:
            wrapped = f'<root xmlns:w="{_W}" xmlns:r="{_R}">{edit.content}</root>'
            root = etree.fromstring(wrapped.encode("utf-8"))
            for elem in root.iter():
                if callable(elem.tag):
                    continue
                for attr in [f"{{{_R}}}id", f"{{{_R}}}embed", f"{{{_R}}}link"]:
                    val = elem.get(attr)
                    if val and val not in valid_ids:
                        click.echo(
                            f"  Warning: section '{edit.target_heading}' references "
                            f"unknown relationship ID '{val}' — may produce broken link in output."
                        )
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Option B — Additive bullet patch
# Option C — In-place text correction for single-run paragraphs
# ---------------------------------------------------------------------------

def _is_list_para(elem: etree._Element) -> bool:
    """Return True if this w:p has w:numPr (list/bullet styling)."""
    ppr = elem.find(f"{{{_W}}}pPr")
    if ppr is None:
        return False
    return ppr.find(f"{{{_W}}}numPr") is not None


def _clean_md_inline(text: str) -> str:
    """Strip inline markdown markers, preserving alphanumeric content."""
    return _MD_INLINE_RE.sub("", text).strip()


def _clean_md_table_cell(text: str, *, keep_bold: bool = False) -> str:
    """Clean markdown table cell: unescape backslashes, strip emphasis but keep underscores.

    If *keep_bold* is True, ``**`` markers are preserved so that downstream
    code can create bold/plain run boundaries.
    """
    val = re.sub(r"\\([_*\\`~|])", r"\1", text)
    if keep_bold:
        # Strip everything except ** bold markers
        return re.sub(r"`{1,3}|~~", "", val).strip()
    return _MD_TABLE_CELL_RE.sub("", val).strip()


def _set_bullet_text_with_formatting(
    para: etree._Element, md_text_raw: str, template_run: etree._Element
) -> None:
    """Set paragraph text from raw MD, preserving **bold** boundaries as separate runs.

    Input:  '**Test Strategy & Framework:** Define the test approach...'
    Result: [Run(bold, "Test Strategy & Framework:"), Run(normal, " Define the test approach...")]
    """
    # Remove all existing w:r elements from para
    for old_r in para.findall(f"{{{_W}}}r"):
        para.remove(old_r)

    # Split on **...** boundaries
    parts = re.split(r"(\*\*[^*]+?\*\*)", md_text_raw)

    for part in parts:
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**")
        text = part[2:-2] if is_bold else part

        new_run = copy.deepcopy(template_run)
        rpr = new_run.find(f"{{{_W}}}rPr")
        if rpr is None:
            rpr = etree.SubElement(new_run, f"{{{_W}}}rPr")

        bold_elem = rpr.find(f"{{{_W}}}b")
        if is_bold and bold_elem is None:
            etree.SubElement(rpr, f"{{{_W}}}b")
        elif not is_bold and bold_elem is not None:
            rpr.remove(bold_elem)
        # Also handle w:bCs (bold complex script) for consistency
        bcs = rpr.find(f"{{{_W}}}bCs")
        if is_bold and bcs is None:
            etree.SubElement(rpr, f"{{{_W}}}bCs")
        elif not is_bold and bcs is not None:
            rpr.remove(bcs)

        # Set text on first w:t, remove extras from cloned run
        t_elems = new_run.findall(f"{{{_W}}}t")
        if t_elems:
            t_elems[0].text = text
            t_elems[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for extra_t in t_elems[1:]:
                new_run.remove(extra_t)
        else:
            t_elem = etree.SubElement(new_run, f"{{{_W}}}t")
            t_elem.text = text
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        para.append(new_run)


def _patch_new_bullets(
    body: etree._Element,
    section_elements: list[etree._Element],
    insert_pos: int,
    md_content: str,
) -> int:
    """Option B: insert/update MD bullets in the DOCX section.

    Similarity bands:
      >= 0.80  — already present, skip
      0.55-0.80 — similar but changed, update in place (Fix 2)
      < 0.55  — genuinely new, insert at neighbor-anchored position (Fix 1)

    Uses _set_bullet_text_with_formatting for inline bold preservation (Fix 3).
    Returns the number of bullets inserted or updated.
    """
    from difflib import SequenceMatcher

    # Track both raw (with **) and cleaned bullet text
    md_bullet_pairs: list[tuple[str, str]] = []  # (cleaned, raw)
    for m in _MD_BULLET_RE.finditer(md_content):
        raw = m.group(1).strip()
        cleaned = _clean_md_inline(raw)
        if cleaned:
            md_bullet_pairs.append((cleaned, raw))
    if not md_bullet_pairs:
        return 0

    list_paras = [e for e in section_elements if e.tag == f"{{{_W}}}p" and _is_list_para(e)]
    template = list_paras[-1] if list_paras else None
    if template is None:
        return 0

    # Get template run for formatting helper
    template_runs = template.findall(f"{{{_W}}}r")
    template_run = template_runs[0] if template_runs else None
    if template_run is None:
        return 0

    # 1. Build ordered anchors: (cleaned, raw, best_match_para, ratio)
    md_bullet_anchors: list[tuple[str, str, etree._Element | None, float]] = []
    for cleaned, raw in md_bullet_pairs:
        best_match_para = None
        best_ratio = 0.0
        for p in list_paras:
            docx_text = _clean_md_inline(
                "".join(t.text or "" for t in p.findall(f".//{{{_W}}}t")).strip()
            )
            ratio = SequenceMatcher(
                None, cleaned.lower(), docx_text.lower(), autojunk=False
            ).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match_para = p if ratio >= 0.55 else None
        md_bullet_anchors.append((cleaned, raw, best_match_para, best_ratio))

    # 2. Process each MD bullet based on similarity band
    changed = 0
    for i, (cleaned, raw, match_para, ratio) in enumerate(md_bullet_anchors):
        if ratio >= 0.80:
            # If MD version is substantially longer (>20% more text),
            # update the DOCX bullet even though it's a close match —
            # the extra content is meaningful (e.g., trailing sentence).
            if match_para is not None and len(cleaned) > len(
                _clean_md_inline(
                    "".join(
                        t.text or ""
                        for t in match_para.findall(f".//{{{_W}}}t")
                    ).strip()
                )
            ) * 1.10:
                _set_bullet_text_with_formatting(
                    match_para, raw, template_run
                )
                changed += 1
            continue

        if 0.55 <= ratio < 0.80 and match_para is not None:
            # Fix 2: Update existing bullet in place
            _set_bullet_text_with_formatting(match_para, raw, template_run)
            changed += 1
            continue

        # ratio < 0.55: Genuinely new — find insertion point.
        # But first, skip if the bullet text already appears as a heading or
        # body paragraph in the DOCX section (e.g. "Client Reference" exists
        # as a heading but MD has it as "1. Client Reference" list item).
        _skip_insert = False
        for _se in section_elements:
            if _se.tag != f"{{{_W}}}p":
                continue
            _se_text = "".join(
                t.text or "" for t in _se.findall(f".//{{{_W}}}t")
            ).strip().lower()
            if _se_text and cleaned.lower() in _se_text:
                _skip_insert = True
                break
        if _skip_insert:
            continue

        # First, try to find the preceding sub-heading in the MD and
        # locate it in the DOCX body to maintain sub-section placement.
        anchor_para = None

        # Find which MD line this bullet is on, then find the nearest
        # preceding non-bullet, non-blank line (sub-heading candidate).
        md_lines = md_content.splitlines()
        bullet_line_idx = None
        for li, ln in enumerate(md_lines):
            if _clean_md_inline(ln.strip().lstrip("*- ")) == cleaned:
                bullet_line_idx = li
                break
            # Also try prefix match for long bullets
            if cleaned[:40] and _clean_md_inline(ln.strip().lstrip("*- ")).startswith(cleaned[:40]):
                bullet_line_idx = li
                break

        if bullet_line_idx is not None:
            # Walk backwards to find sub-heading
            for li in range(bullet_line_idx - 1, -1, -1):
                ln = md_lines[li].strip()
                if not ln:
                    continue
                if _MD_BULLET_RE.match(md_lines[li]):
                    continue
                if ln.startswith("|") or ln.startswith("#"):
                    continue
                # This is a sub-heading candidate (plain text line)
                sub_heading_text = _clean_md_inline(ln).lower()
                if len(sub_heading_text) < 3:
                    continue
                # Count which occurrence of this text precedes the bullet in the MD
                _md_occurrence = 0
                for _mli in range(li + 1):
                    _mln = _clean_md_inline(md_lines[_mli].strip()).lower()
                    if _mln == sub_heading_text or (
                        len(sub_heading_text) > 10 and sub_heading_text in _mln
                    ):
                        _md_occurrence += 1
                # Find the same-numbered occurrence in DOCX body
                body_children = list(body)
                _docx_occurrence = 0
                for bi, elem in enumerate(body_children):
                    if elem.tag != f"{{{_W}}}p":
                        continue
                    if elem not in section_elements:
                        continue
                    para_text = "".join(
                        t.text or "" for t in elem.findall(f".//{{{_W}}}t")
                    ).strip().lower()
                    if para_text == sub_heading_text or (
                        len(sub_heading_text) > 10
                        and sub_heading_text in para_text
                    ):
                        _docx_occurrence += 1
                        if _docx_occurrence < _md_occurrence:
                            continue
                        # Found the sub-heading in DOCX. Insert after
                        # the last list paragraph following it.
                        last_list_after = bi
                        for ki in range(bi + 1, len(body_children)):
                            if body_children[ki] not in section_elements:
                                break
                            if _is_list_para(body_children[ki]):
                                last_list_after = ki
                            elif body_children[ki].tag == f"{{{_W}}}p":
                                # Non-list paragraph after sub-heading's
                                # list — stop here
                                p_text = "".join(
                                    t.text or ""
                                    for t in body_children[ki].findall(
                                        f".//{{{_W}}}t"
                                    )
                                ).strip()
                                if p_text and not _is_list_para(
                                    body_children[ki]
                                ):
                                    break
                        anchor_para = body_children[last_list_after]
                        break
                if anchor_para is not None:
                    break

        # Fall back to neighbor-anchored insertion
        if anchor_para is None:
            for j in range(i - 1, -1, -1):
                _, _, prev_match, prev_ratio = md_bullet_anchors[j]
                if prev_match is not None and prev_ratio >= 0.55:
                    anchor_para = prev_match
                    break

        new_para = copy.deepcopy(template)
        _set_bullet_text_with_formatting(new_para, raw, template_run)

        if anchor_para is not None:
            anchor_idx = list(body).index(anchor_para)
            body.insert(anchor_idx + 1, new_para)
        elif list_paras:
            body.insert(list(body).index(list_paras[0]), new_para)
        else:
            body.insert(insert_pos, new_para)

        changed += 1

    return changed


def _remove_stale_bullets(
    body: etree._Element,
    section_elements: list[etree._Element],
    md_content: str,
) -> int:
    """Option E: remove DOCX bullets that have no match in the source MD.

    After Option B has added/updated bullets, any DOCX list paragraph whose
    best match against ALL source MD bullets is below 0.50 is considered stale
    (the user removed it from the MD) and is deleted.

    Returns the number of bullets removed.
    """
    from difflib import SequenceMatcher

    # Gather cleaned MD bullet texts
    md_bullets: list[str] = []
    for m in _MD_BULLET_RE.finditer(md_content):
        cleaned = _clean_md_inline(m.group(1).strip())
        if cleaned:
            md_bullets.append(cleaned.lower())
    if not md_bullets:
        return 0

    list_paras = [
        e for e in section_elements
        if e.tag == f"{{{_W}}}p" and _is_list_para(e)
    ]
    if not list_paras:
        return 0

    stale: list[etree._Element] = []
    for p in list_paras:
        docx_text = _clean_md_inline(
            "".join(t.text or "" for t in p.findall(f".//{{{_W}}}t")).strip()
        ).lower()
        if not docx_text:
            continue
        best = max(
            SequenceMatcher(None, docx_text, mb, autojunk=False).ratio()
            for mb in md_bullets
        )
        if best < 0.50:
            stale.append(p)

    for p in stale:
        body.remove(p)

    return len(stale)


def _iter_section_paragraphs(section_elements: list[etree._Element]):
    """Yield all w:p elements in section, including those inside table cells."""
    for elem in section_elements:
        if elem.tag == f"{{{_W}}}p":
            yield elem
        elif elem.tag == f"{{{_W}}}tbl":
            for cell_p in elem.findall(f".//{{{_W}}}p"):
                yield cell_p


def _parse_md_table_blocks(
    md_content: str, *, keep_bold: bool = False
) -> list[list[list[str]]]:
    """Parse markdown content into separate table blocks.

    Each block is a list of rows (each row a list of cell values).
    Tables are separated by non-table lines (paragraphs, headings, blank lines).

    If *keep_bold* is True, ``**`` markers are preserved in cell values so
    that downstream code can create bold/plain run boundaries.
    """
    blocks: list[list[list[str]]] = []
    current_block: list[list[str]] = []

    for line in md_content.splitlines():
        stripped = line.strip()
        if "|" in stripped and not stripped.startswith("#"):
            # Split on pipe, strip first/last empty entries from | delimiters,
            # but preserve interior empty cells to maintain column alignment.
            raw_cells = stripped.split("|")
            # Leading/trailing pipes produce empty strings at edges
            if raw_cells and not raw_cells[0].strip():
                raw_cells = raw_cells[1:]
            if raw_cells and not raw_cells[-1].strip():
                raw_cells = raw_cells[:-1]
            cells = [
                _clean_md_table_cell(c.strip(), keep_bold=keep_bold)
                for c in raw_cells
            ]
            if not cells:
                continue
            # Skip separator rows (all dashes/colons)
            if all(set(c) <= set("-: ") for c in cells):
                continue
            current_block.append(cells)
        else:
            if current_block:
                blocks.append(current_block)
                current_block = []
    if current_block:
        blocks.append(current_block)
    return blocks


def _extract_tr_cells(tr: etree._Element) -> list[str]:
    """Extract text content from each w:tc in a w:tr."""
    cells: list[str] = []
    for tc in tr.findall(f"{{{_W}}}tc"):
        cell_text = " ".join(
            (t.text or "") for t in tc.findall(f".//{{{_W}}}t")
        ).strip()
        cells.append(cell_text)
    return cells


def _parse_bold_segments(text: str) -> list[tuple[str, bool]]:
    """Parse markdown bold markers and return [(text, is_bold), ...].

    Handles **bold** markers.  Returns plain text if no markers found.
    """
    segments: list[tuple[str, bool]] = []
    pos = 0
    while pos < len(text):
        start = text.find("**", pos)
        if start == -1:
            segments.append((text[pos:], False))
            break
        if start > pos:
            segments.append((text[pos:start], False))
        end = text.find("**", start + 2)
        if end == -1:
            # Unclosed marker — treat rest as plain
            segments.append((text[start:], False))
            break
        segments.append((text[start + 2:end], True))
        pos = end + 2
    return [(t, b) for t, b in segments if t]


def _set_tr_cell_texts(tr: etree._Element, md_cells: list[str]) -> int:
    """Update cell text values in a w:tr from MD cell values.

    For each w:tc, finds the first w:p with a w:r and updates its w:t.
    Preserves bold run boundaries when MD cells contain **bold** markers.
    Returns the number of cells changed.
    """
    tcs = tr.findall(f"{{{_W}}}tc")
    changed = 0
    for i, tc in enumerate(tcs):
        if i >= len(md_cells):
            break
        md_val = md_cells[i]
        # Get current cell text (strip markdown bold for comparison)
        current = " ".join(
            (t.text or "") for t in tc.findall(f".//{{{_W}}}t")
        ).strip()
        plain_md = md_val.replace("**", "")
        if current == plain_md:
            continue
        # Find the first paragraph with a run to update
        all_paras = tc.findall(f"{{{_W}}}p")
        updated = False
        for p in all_paras:
            runs = p.findall(f"{{{_W}}}r")
            if not runs:
                continue

            # Parse bold segments from MD content
            segments = _parse_bold_segments(md_val)
            has_bold_markers = any(b for _, b in segments)

            if has_bold_markers and len(segments) > 1:
                # Build new runs matching bold boundaries.
                # Preserve pPr (paragraph properties) from the original paragraph.
                # Use first run as template for run properties.
                template_rpr = runs[0].find(f"{{{_W}}}rPr")

                # Remove all existing runs
                for r in runs:
                    p.remove(r)

                for seg_text, seg_bold in segments:
                    new_run = etree.SubElement(p, f"{{{_W}}}r")
                    # Clone base run properties (font, size, etc.) from template
                    if template_rpr is not None:
                        new_rpr = copy.deepcopy(template_rpr)
                    else:
                        new_rpr = etree.SubElement(new_run, f"{{{_W}}}rPr")
                    # Set or remove bold
                    b_elem = new_rpr.find(f"{{{_W}}}b")
                    bcs_elem = new_rpr.find(f"{{{_W}}}bCs")
                    if seg_bold:
                        if b_elem is None:
                            etree.SubElement(new_rpr, f"{{{_W}}}b")
                        if bcs_elem is None:
                            etree.SubElement(new_rpr, f"{{{_W}}}bCs")
                    else:
                        if b_elem is not None:
                            new_rpr.remove(b_elem)
                        if bcs_elem is not None:
                            new_rpr.remove(bcs_elem)
                    # Only add rPr if it was created from template or has children
                    if template_rpr is not None:
                        new_run.insert(0, new_rpr)
                    elif len(new_rpr) > 0:
                        new_run.insert(0, new_rpr)
                    t_elem = etree.SubElement(new_run, f"{{{_W}}}t")
                    t_elem.text = seg_text
                    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            else:
                # No bold markers — original single-run behavior
                plain_text = plain_md
                t_elems = runs[0].findall(f"{{{_W}}}t")
                if t_elems:
                    t_elems[0].text = plain_text
                    t_elems[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for extra_t in t_elems[1:]:
                        runs[0].remove(extra_t)
                else:
                    t_elem = etree.SubElement(runs[0], f"{{{_W}}}t")
                    t_elem.text = plain_text
                    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                # Remove extra runs to avoid duplicate text
                for extra_r in runs[1:]:
                    p.remove(extra_r)
            updated = True
            break
        # Remove extra paragraphs in the cell to prevent duplication
        # (multi-line cells have multiple w:p; we set the full text on
        # the first one, so the rest would duplicate content).
        if updated and len(all_paras) > 1:
            for extra_p in all_paras[1:]:
                tc.remove(extra_p)
        if updated:
            changed += 1
    return changed


def _patch_table_rows(
    section_elements: list[etree._Element],
    md_content: str,
    body: etree._Element,
) -> tuple[int, int, int]:
    """Option D: targeted table row patching for unchanged sections.

    For each w:tbl in the section:
    1. Find the best-matching MD table block (by overall text similarity)
    2. Match DOCX rows to MD rows within that block
    3. Update cells in matched rows that differ
    4. Insert new MD rows (no DOCX match) by cloning a template row
    5. Remove stale DOCX rows that share a key-column value with an inserted row

    Returns (rows_updated, rows_inserted, rows_removed).
    """
    from difflib import SequenceMatcher

    md_blocks = _parse_md_table_blocks(md_content)
    md_blocks_raw = _parse_md_table_blocks(md_content, keep_bold=True)
    if not md_blocks:
        return 0, 0, 0

    total_updated = 0
    total_inserted = 0
    total_removed = 0

    # Collect all DOCX tables in the section
    docx_tables = [e for e in section_elements if e.tag == f"{{{_W}}}tbl"]
    if not docx_tables:
        return 0, 0, 0

    # Phase 0: Match each DOCX table to the best MD table block.
    # This prevents cross-table contamination when a section has many sub-tables.
    md_block_used: set[int] = set()
    # Each entry: (tbl, cleaned_rows, raw_rows_with_bold_markers)
    table_block_pairs: list[
        tuple[etree._Element, list[list[str]], list[list[str]]]
    ] = []

    for tbl in docx_tables:
        docx_trs = tbl.findall(f"{{{_W}}}tr")
        if not docx_trs:
            continue
        # Build aggregate text for the DOCX table
        docx_texts = []
        for tr in docx_trs:
            cells = _extract_tr_cells(tr)
            docx_texts.append(" ".join(cells))
        docx_full = " ".join(docx_texts).lower().strip()
        if not docx_full:
            continue

        # Find the best matching MD block
        best_bi = -1
        best_ratio = 0.0
        for bi, block in enumerate(md_blocks):
            if bi in md_block_used:
                continue
            md_full = " ".join(" ".join(row) for row in block).lower().strip()
            if not md_full:
                continue
            ratio = SequenceMatcher(
                None, docx_full, md_full, autojunk=False
            ).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_bi = bi

        if best_bi >= 0 and best_ratio >= 0.20:
            # Skip small structural tables (signature blocks, header tables)
            # when the match quality is low — prevents corruption.
            docx_data_rows = len(docx_trs) - 1  # exclude header
            if docx_data_rows <= 4 and best_ratio < 0.50:
                continue
            # Skip tables with inconsistent column counts across rows
            # (merged cells) — _set_tr_cell_texts can't handle these safely.
            col_counts = set()
            for _tr in docx_trs:
                col_counts.add(len(_tr.findall(f"{{{_W}}}tc")))
            if len(col_counts) > 1:
                continue
            md_block_used.add(best_bi)
            table_block_pairs.append(
                (tbl, md_blocks[best_bi], md_blocks_raw[best_bi])
            )

    # Phase 1-3: For each matched table-block pair, do row-level matching
    for tbl, md_table_rows, md_table_rows_raw in table_block_pairs:
        docx_trs = tbl.findall(f"{{{_W}}}tr")
        if not docx_trs:
            continue

        # Build DOCX row texts for matching
        docx_row_data: list[tuple[etree._Element, list[str], str]] = []
        for tr in docx_trs:
            cells = _extract_tr_cells(tr)
            row_text = " ".join(cells).lower().strip()
            docx_row_data.append((tr, cells, row_text))

        # Track which MD rows are consumed (matched to a DOCX row)
        md_matched: list[bool] = [False] * len(md_table_rows)
        docx_matched_md_idx: dict[int, int] = {}  # docx_idx -> md_idx

        # Build a match matrix, then greedily assign best matches
        matches: list[tuple[float, int, int]] = []  # (ratio, md_idx, docx_idx)
        for mi, md_row in enumerate(md_table_rows):
            md_text = " ".join(md_row).lower().strip()
            if not md_text:
                continue
            for di, (_, _, docx_text) in enumerate(docx_row_data):
                if not docx_text:
                    continue
                ratio = SequenceMatcher(
                    None, md_text, docx_text, autojunk=False
                ).ratio()
                if ratio >= 0.50:
                    matches.append((ratio, mi, di))

        # Sort by similarity descending, greedily assign
        matches.sort(key=lambda x: -x[0])
        docx_used: set[int] = set()
        row_assignments: list[tuple[int, int, float]] = []
        for ratio, mi, di in matches:
            if md_matched[mi] or di in docx_used:
                continue
            md_matched[mi] = True
            docx_used.add(di)
            docx_matched_md_idx[di] = mi
            row_assignments.append((mi, di, ratio))

        # Update matched rows where cells differ
        for mi, di, ratio in row_assignments:
            md_row = md_table_rows[mi]
            tr, docx_cells, _ = docx_row_data[di]
            needs_update = False
            for ci in range(min(len(md_row), len(docx_cells))):
                if md_row[ci] != docx_cells[ci]:
                    needs_update = True
                    break
            if not needs_update:
                continue
            # Pass raw cells (with **bold** markers) so _set_tr_cell_texts
            # can create proper bold/plain run boundaries.
            changed = _set_tr_cell_texts(tr, md_table_rows_raw[mi])
            if changed:
                total_updated += 1

        # Save original row count before any inserts (for Phase 4 Case B)
        original_docx_data_count = len(docx_trs) - 1  # exclude header

        # Insert unmatched MD rows (new rows)
        unmatched_md = [i for i, matched in enumerate(md_matched) if not matched]
        inserted_md_indices: set[int] = set()
        skip_insert = False

        if unmatched_md:
            # Safety guard: don't insert if unmatched count is disproportionately large.
            docx_row_count = len(docx_trs)
            if len(unmatched_md) > max(docx_row_count, 5):
                click.echo(
                    f"    Skipping {len(unmatched_md)} row insertion(s) — "
                    f"exceeds DOCX table size ({docx_row_count} rows)"
                )
                skip_insert = True

            if not skip_insert:
                # Find a template row (last data row — skip header row at index 0)
                template_tr = docx_trs[-1] if len(docx_trs) > 1 else docx_trs[0]

                for mi in unmatched_md:
                    md_row = md_table_rows[mi]
                    if not any(md_row):
                        continue

                    new_tr = copy.deepcopy(template_tr)
                    tcs = new_tr.findall(f"{{{_W}}}tc")

                    # Ensure new row has enough cells — clone last cell if needed
                    while len(tcs) < len(md_row):
                        new_tc = copy.deepcopy(tcs[-1])
                        new_tr.append(new_tc)
                        tcs = new_tr.findall(f"{{{_W}}}tc")

                    # Pass raw cells (with **bold** markers) for proper formatting
                    _set_tr_cell_texts(new_tr, md_table_rows_raw[mi])

                    # Find insertion position: after the nearest preceding matched DOCX row
                    anchor_di = None
                    for prev_mi in range(mi - 1, -1, -1):
                        for di, mapped_mi in docx_matched_md_idx.items():
                            if mapped_mi == prev_mi:
                                anchor_di = di
                                break
                        if anchor_di is not None:
                            break

                    if anchor_di is not None:
                        # Use element ref from docx_row_data (stable) instead
                        # of docx_trs (refreshed after each insert, indices shift).
                        anchor_tr = docx_row_data[anchor_di][0]
                        anchor_idx = list(tbl).index(anchor_tr)
                        tbl.insert(anchor_idx + 1, new_tr)
                    else:
                        tbl.append(new_tr)

                    total_inserted += 1
                    inserted_md_indices.add(mi)
                    docx_trs = tbl.findall(f"{{{_W}}}tr")
                    docx_matched_md_idx[len(docx_row_data)] = mi
                    docx_row_data.append((new_tr, md_row, " ".join(md_row).lower()))

        # Phase 4: Remove stale rows.
        # Case A: DOCX row unmatched + shares a key value (col 1 or col 2)
        #         with a newly inserted row → it was replaced.
        # Case B: MD table has fewer data rows than DOCX table, and unmatched
        #         DOCX rows have no close match in MD → pure deletion.
        stale_trs: list[etree._Element] = []

        # Collect key values (columns 1 and 2) of inserted rows for Case A
        inserted_keys: set[str] = set()
        for mi in inserted_md_indices:
            for cell in md_table_rows[mi][:2]:  # check first 2 columns
                val = cell.strip()
                if val:
                    inserted_keys.add(val.lower())

        for di, (tr, cells, row_text) in enumerate(docx_row_data):
            if di in docx_matched_md_idx:
                continue  # matched or inserted — not stale
            if di == 0:
                continue  # skip header row

            # Case A: key-column overlap with an inserted row
            if inserted_keys:
                is_key_overlap = False
                for cell in cells[:2]:
                    val = cell.strip()
                    if val and val.lower() in inserted_keys:
                        is_key_overlap = True
                        break
                if is_key_overlap:
                    stale_trs.append(tr)
                    continue

            # Case B: all MD rows found matches but this DOCX row didn't →
            # it was deleted in v2 (surplus row with no MD counterpart).
            if all(md_matched) and row_text:
                stale_trs.append(tr)

        # Case C: row-count balance. After inserts and Case A/B removals,
        # if the table would have more data rows than the MD table, remove
        # remaining unmatched DOCX rows to balance.
        md_data_count = len(md_table_rows) - 1  # exclude header
        final_data_count = (
            original_docx_data_count
            + len(inserted_md_indices)
            - len(stale_trs)
        )
        if final_data_count > md_data_count:
            surplus = final_data_count - md_data_count
            for di, (tr, cells, row_text) in enumerate(docx_row_data):
                if surplus <= 0:
                    break
                if di in docx_matched_md_idx:
                    continue
                if di == 0:
                    continue
                if tr in stale_trs:
                    continue
                if row_text:
                    stale_trs.append(tr)
                    surplus -= 1

        for tr in stale_trs:
            tbl.remove(tr)
            total_removed += 1

    return total_updated, total_inserted, total_removed


def _inject_bullet_styles(
    document_xml_bytes: bytes,
    mappings: list["SectionMapping"],
    docx_path: Path | None = None,
) -> tuple[bytes, int]:
    """Post-LLM fix: inject ListParagraph style on unstyled paragraphs that
    correspond to MD bullet lines (``* text`` or ``- text``).

    The LLM often generates bullet content as plain ``<w:p>`` without the
    ``ListParagraph`` style or ``<w:numPr>`` needed for Word to render bullets.
    This function finds a bullet-style template from an existing paragraph in
    the document and clones its ``pPr`` onto matching unstyled paragraphs.

    Returns (modified_xml_bytes, paragraphs_styled).
    """
    root = etree.fromstring(document_xml_bytes)
    body = root.find(f"{{{_W}}}body")
    if body is None:
        return document_xml_bytes, 0

    # Build set of bullet numIds from numbering.xml so we pick a bullet
    # template rather than a numbered-list template.
    bullet_num_ids: set[str] = set()
    if docx_path is not None:
        try:
            with zipfile.ZipFile(docx_path, "r") as zf:
                if "word/numbering.xml" in zf.namelist():
                    num_root = etree.fromstring(zf.read("word/numbering.xml"))
                    # Map abstractNumId -> numFmt at level 0
                    abstract_fmt: dict[str, str] = {}
                    for abstract in num_root.findall(f"{{{_W}}}abstractNum"):
                        aid = abstract.get(f"{{{_W}}}abstractNumId", "")
                        lvl0 = abstract.find(
                            f"{{{_W}}}lvl[@{{{_W}}}ilvl='0']"
                        )
                        if lvl0 is not None:
                            fmt_el = lvl0.find(f"{{{_W}}}numFmt")
                            if fmt_el is not None:
                                abstract_fmt[aid] = fmt_el.get(
                                    f"{{{_W}}}val", ""
                                )
                    # Map numId -> abstractNumId, keep only bullets
                    for num in num_root.findall(f"{{{_W}}}num"):
                        nid = num.get(f"{{{_W}}}numId", "")
                        abs_ref = num.find(f"{{{_W}}}abstractNumId")
                        if abs_ref is not None:
                            aid = abs_ref.get(f"{{{_W}}}val", "")
                            if abstract_fmt.get(aid) == "bullet":
                                bullet_num_ids.add(nid)
        except Exception:
            pass  # fall back to accepting any numId

    # Find a template pPr from an existing ListParagraph with numPr
    template_ppr: etree._Element | None = None
    for p in body.findall(f".//{{{_W}}}p"):
        ppr = p.find(f"{{{_W}}}pPr")
        if ppr is None:
            continue
        style_el = ppr.find(f"{{{_W}}}pStyle")
        if style_el is not None and style_el.get(f"{{{_W}}}val") == "ListParagraph":
            num_pr = ppr.find(f"{{{_W}}}numPr")
            if num_pr is not None:
                # If we know which numIds are bullets, filter
                if bullet_num_ids:
                    nid_el = num_pr.find(f"{{{_W}}}numId")
                    nid = nid_el.get(f"{{{_W}}}val", "") if nid_el is not None else ""
                    if nid not in bullet_num_ids:
                        continue
                template_ppr = copy.deepcopy(ppr)
                # Remove any run-level or spacing properties that are
                # paragraph-specific — keep only style + numPr
                for child in list(template_ppr):
                    tag = etree.QName(child).localname
                    if tag not in ("pStyle", "numPr"):
                        template_ppr.remove(child)
                break

    if template_ppr is None:
        return document_xml_bytes, 0

    total_styled = 0

    for mapping in mappings:
        if mapping.action != "replace":
            continue

        # Extract bullet lines from MD content
        bullet_texts: list[str] = []
        for line in mapping.md_content.splitlines():
            stripped = line.strip()
            if stripped.startswith("* ") or stripped.startswith("- "):
                # Normalize: strip marker, inline markdown, extra whitespace
                text = stripped[2:].strip()
                text = re.sub(r"[*_`~]", "", text).strip()
                if text:
                    bullet_texts.append(text.lower())

        if not bullet_texts:
            continue

        heading = mapping.md_heading
        try:
            start, end = _find_section_range(body, heading)
        except ValueError:
            continue

        children = list(body)[start:end]
        for p in children:
            if p.tag != f"{{{_W}}}p":
                continue
            ppr = p.find(f"{{{_W}}}pPr")

            # Case 1: ListParagraph without numPr — LLM used the style but
            # omitted the numbering reference, so Word won't render bullets.
            if ppr is not None:
                style_el = ppr.find(f"{{{_W}}}pStyle")
                if style_el is not None:
                    style_val = style_el.get(f"{{{_W}}}val")
                    if style_val == "ListParagraph" and ppr.find(f"{{{_W}}}numPr") is None:
                        # Inject numPr from template
                        num_pr = template_ppr.find(f"{{{_W}}}numPr")
                        if num_pr is not None:
                            ppr.insert(1, copy.deepcopy(num_pr))  # after pStyle
                            total_styled += 1
                    continue  # skip further checks for styled paras

            # Case 2: Unstyled paragraph matching MD bullet text
            para_text = "".join(
                t.text or "" for t in p.findall(f".//{{{_W}}}t")
            ).strip().lower()
            if not para_text:
                continue

            matched = False
            for bt in bullet_texts:
                if para_text.startswith(bt[:40]) or bt.startswith(para_text[:40]):
                    matched = True
                    break

            if not matched:
                continue

            new_ppr = copy.deepcopy(template_ppr)
            if ppr is not None:
                rpr = ppr.find(f"{{{_W}}}rPr")
                if rpr is not None:
                    new_ppr.append(copy.deepcopy(rpr))
                p.remove(ppr)
            p.insert(0, new_ppr)
            total_styled += 1

    if total_styled:
        return (
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
            total_styled,
        )
    return document_xml_bytes, 0


def _apply_opcodes_to_run(
    run_text: str,
    full_old: str,
    full_new: str,
    opcodes: list[tuple[str, int, int, int, int]],
    run_start: int,
) -> str:
    """Apply SequenceMatcher opcodes to a single run's text.

    Given opcodes that transform full_old → full_new, extract the changes
    that fall within this run (starting at run_start in full_old) and
    produce the updated run text.
    """
    run_end = run_start + len(run_text)
    result = []
    pos = 0  # position within run_text

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "insert":
            # Inserts (i1==i2): attribute to the run containing position i1,
            # using half-open interval [run_start, run_end).  Boundary inserts
            # at run_end go to the next run (whose run_start == i1).
            if not (run_start <= i1 < run_end):
                continue
        else:
            # Skip opcodes entirely before this run
            if i2 <= run_start:
                continue
            # Stop at opcodes entirely after this run
            if i1 >= run_end:
                break

        # Clamp to run boundaries
        eff_i1 = max(i1, run_start)
        eff_i2 = min(i2, run_end)
        local_i1 = eff_i1 - run_start
        local_i2 = eff_i2 - run_start

        # Add any unchanged text before this opcode within the run
        if local_i1 > pos:
            result.append(run_text[pos:local_i1])

        if tag == "equal":
            result.append(run_text[local_i1:local_i2])
        elif tag == "replace":
            # Proportionally map replacement text
            if i2 > i1:
                frac_start = (eff_i1 - i1) / (i2 - i1)
                frac_end = (eff_i2 - i1) / (i2 - i1)
            else:
                frac_start, frac_end = 0.0, 1.0
            new_text = full_new[j1:j2]
            ns = int(frac_start * len(new_text))
            ne = int(frac_end * len(new_text))
            result.append(new_text[ns:ne])
        elif tag == "insert":
            # Insert happens at i1 in old text; if within our run, add it
            if run_start <= i1 <= run_end:
                result.append(full_new[j1:j2])
            result.append(run_text[local_i1:local_i2])
        elif tag == "delete":
            pass  # skip deleted text

        pos = local_i2

    # Add any remaining run text after last opcode
    if pos < len(run_text):
        result.append(run_text[pos:])

    return "".join(result)


def _patch_text_corrections(
    section_elements: list[etree._Element],
    md_content: str,
    min_ratio: float = 0.75,
) -> int:
    """Option C: update w:t text in-place for paragraphs that differ from MD.

    For uniform-formatting paragraphs, rewrites text into the first run.
    For heterogeneous-formatting paragraphs (e.g. bold label + plain text),
    applies targeted character-level patches to preserve formatting.
    List paragraphs are matched against MD bullet lines; non-list paragraphs
    against plain MD lines. Table corrections are handled by Option D.
    *min_ratio* controls the similarity threshold (0.75 for unchanged sections,
    0.90 for LLM-edited sections where only near-exact mismatches should be fixed).
    Returns the number of paragraphs corrected.
    """
    from difflib import SequenceMatcher

    md_lines = []
    md_bullet_lines = []
    for line in md_content.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        bm = _MD_BULLET_RE.match(line)
        if bm:
            md_bullet_lines.append(_clean_md_inline(bm.group(1).strip()))
        else:
            md_lines.append(_clean_md_inline(stripped))

    if not md_lines and not md_bullet_lines:
        return 0

    corrected = 0

    # --- Non-table paragraph corrections (tables handled by Option D) ---
    for elem in section_elements:
        if elem.tag != f"{{{_W}}}p":
            continue
        if _para_heading_text(elem) is not None:
            continue
        # Skip paragraphs containing structured document tags (w:sdt) —
        # these are content controls whose text is not in normal w:r runs.
        # Correcting run text without accounting for SDT content causes
        # duplication (the SDT retains its text and the run gets a copy).
        if elem.find(f".//{{{_W}}}sdt") is not None:
            continue
        is_list = _is_list_para(elem)
        pool = md_bullet_lines if is_list else md_lines
        if not pool:
            continue
        runs = elem.findall(f"{{{_W}}}r")
        if not runs:
            continue
        # For multi-run paragraphs, check if all runs share the same
        # formatting (no inline bold/italic differences).
        heterogeneous = False
        if len(runs) > 1:
            rpr_sigs = set()
            for r in runs:
                rpr = r.find(f"{{{_W}}}rPr")
                if rpr is None:
                    rpr_sigs.add(("none",))
                else:
                    b = rpr.find(f"{{{_W}}}b")
                    i = rpr.find(f"{{{_W}}}i")
                    u = rpr.find(f"{{{_W}}}u")
                    rpr_sigs.add((
                        b is not None,
                        i is not None,
                        u is not None,
                    ))
            if len(rpr_sigs) > 1:
                heterogeneous = True
        # Gather text from ALL runs
        current_text = "".join(
            (t.text or "")
            for r in runs
            for t in r.findall(f"{{{_W}}}t")
        ).strip()
        if not current_text:
            continue

        best_ratio = 0.0
        best_line = None
        for md_line in pool:
            ratio = SequenceMatcher(None, current_text.lower(), md_line.lower(), autojunk=False).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_line = md_line

        if best_line is None or best_ratio >= 1.0 or best_line == current_text:
            continue
        if best_ratio < min_ratio:
            continue
        # Guard: skip if MD line is much longer than DOCX text
        if len(best_line) > len(current_text) * 2:
            continue

        if heterogeneous:
            # For multi-run paragraphs with mixed formatting, apply
            # targeted character-level patches to preserve run formatting.
            # Use raw (unstripped) concatenated text so character positions
            # align exactly with individual run boundaries.
            raw_text = "".join(
                (t.text or "")
                for r in runs
                for t in r.findall(f"{{{_W}}}t")
            )
            # Align best_line with raw_text (re-add stripped whitespace)
            lstrip_n = len(raw_text) - len(raw_text.lstrip())
            rstrip_n = len(raw_text) - len(raw_text.rstrip())
            raw_best = raw_text[:lstrip_n] + best_line + raw_text[len(raw_text) - rstrip_n:] if rstrip_n else raw_text[:lstrip_n] + best_line
            opcodes = SequenceMatcher(
                None, raw_text, raw_best, autojunk=False
            ).get_opcodes()
            # Build per-run text boundaries
            run_texts = []
            for r in runs:
                run_texts.append("".join(
                    t.text or "" for t in r.findall(f"{{{_W}}}t")
                ))
            patched = False
            char_offset = 0
            for ri, rtxt in enumerate(run_texts):
                new_run_text = _apply_opcodes_to_run(
                    rtxt, raw_text, raw_best, opcodes, char_offset
                )
                if new_run_text != rtxt:
                    t_elems_r = runs[ri].findall(f"{{{_W}}}t")
                    if t_elems_r:
                        t_elems_r[0].text = new_run_text
                        t_elems_r[0].set(
                            "{http://www.w3.org/XML/1998/namespace}space",
                            "preserve",
                        )
                        for extra in t_elems_r[1:]:
                            runs[ri].remove(extra)
                    patched = True
                char_offset += len(rtxt)
            if patched:
                corrected += 1
        else:
            t_elems = runs[0].findall(f"{{{_W}}}t")
            if not t_elems:
                continue
            t_elems[0].text = best_line
            for extra_t in t_elems[1:]:
                p = extra_t.getparent()
                if p is not None:
                    p.remove(extra_t)
            # Remove extra runs (multi-run uniform paragraphs)
            for extra_r in runs[1:]:
                elem.remove(extra_r)
            corrected += 1

    return corrected


# ---------------------------------------------------------------------------
# Option C+bold — Apply **bold** emphasis from source MD to DOCX runs
# ---------------------------------------------------------------------------

_BOLD_SPAN_RE = re.compile(r"\*\*(.+?)\*\*")


def _patch_bold_emphasis(
    section_elements: list[etree._Element],
    md_content: str,
) -> int:
    """Apply **bold** markers from source MD to matching DOCX paragraph runs.

    For each DOCX paragraph, finds the best-matching MD line (≥0.75 similarity),
    extracts bold spans from that specific MD line, and applies them using
    word-boundary-aware matching.  This prevents cross-paragraph leakage and
    substring false positives (e.g. "not" inside "notice").

    Returns the number of bold spans applied.
    """
    from difflib import SequenceMatcher

    # Build per-line bold map: for each MD line, extract (cleaned_line, [(phrase, position_ratio)])
    # position_ratio = where in the cleaned line the phrase appears (0.0–1.0)
    md_line_data: list[tuple[str, list[tuple[str, float]]]] = []
    for line in md_content.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        bold_matches = list(_BOLD_SPAN_RE.finditer(stripped))
        if not bold_matches:
            continue
        cleaned = _BOLD_SPAN_RE.sub(lambda m: m.group(1), stripped)
        cleaned_lower = _clean_md_inline(cleaned)
        if not cleaned_lower:
            continue
        phrases_with_pos: list[tuple[str, float]] = []
        # Calculate position by counting how many chars precede each bold match
        # after removing all ** markers
        for bm in bold_matches:
            phrase = bm.group(1).strip()
            if not phrase:
                continue
            # Characters before this match in the original line
            before_match = stripped[:bm.start()]
            # Remove ** markers from the prefix to get true character position
            before_clean = before_match.replace("**", "")
            # Total line length without markers
            total_clean = stripped.replace("**", "")
            ratio = len(before_clean) / max(len(total_clean), 1)
            phrases_with_pos.append((phrase, ratio))
        if phrases_with_pos:
            md_line_data.append((cleaned_lower, phrases_with_pos))

    if not md_line_data:
        return 0

    applied = 0

    for elem in section_elements:
        if elem.tag != f"{{{_W}}}p":
            continue
        if _para_heading_text(elem) is not None:
            continue
        runs = elem.findall(f"{{{_W}}}r")
        if not runs:
            continue

        # Get paragraph text
        full_text_parts: list[str] = []
        run_boundaries: list[tuple[int, int, etree._Element]] = []
        pos = 0
        for r in runs:
            rtxt = "".join(t.text or "" for t in r.findall(f"{{{_W}}}t"))
            run_boundaries.append((pos, pos + len(rtxt), r))
            full_text_parts.append(rtxt)
            pos += len(rtxt)
        full_text = "".join(full_text_parts)
        if not full_text:
            continue

        # Find the best-matching MD line for this paragraph
        para_clean = _clean_md_inline(full_text.strip())
        if not para_clean:
            continue
        best_ratio = 0.0
        best_phrases: list[tuple[str, float]] = []
        for md_clean, phrases_with_pos in md_line_data:
            ratio = SequenceMatcher(
                None, para_clean.lower(), md_clean.lower(), autojunk=False
            ).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_phrases = phrases_with_pos
        if best_ratio < 0.75 or not best_phrases:
            continue

        # Apply each bold phrase from the matched MD line
        para_modified = False
        for phrase, pos_ratio in best_phrases:
            # Word-boundary-aware search in paragraph text
            escaped = re.escape(phrase)
            pattern = re.compile(
                r"(?<![a-zA-Z])" + escaped + r"(?![a-zA-Z])",
                re.IGNORECASE,
            )
            # Find ALL occurrences and pick the one closest to the
            # expected position ratio within the paragraph
            all_matches = list(pattern.finditer(full_text))
            if not all_matches:
                continue
            if len(all_matches) == 1:
                match = all_matches[0]
            else:
                # Pick occurrence closest to the position ratio from the MD line
                best_match = all_matches[0]
                best_dist = float("inf")
                for m in all_matches:
                    m_ratio = m.start() / max(len(full_text), 1)
                    dist = abs(m_ratio - pos_ratio)
                    if dist < best_dist:
                        best_dist = dist
                        best_match = m
                match = best_match

            idx = match.start()
            phrase_end = match.end()

            # Check if already bold
            already_bold = True
            for rstart, rend, relem in run_boundaries:
                if rstart >= phrase_end or rend <= idx:
                    continue
                rpr = relem.find(f"{{{_W}}}rPr")
                if rpr is None or rpr.find(f"{{{_W}}}b") is None:
                    already_bold = False
                    break
            if already_bold:
                continue

            # Split runs at phrase boundaries and apply bold
            new_runs: list[etree._Element] = []
            for rstart, rend, relem in run_boundaries:
                rtxt = "".join(t.text or "" for t in relem.findall(f"{{{_W}}}t"))
                bold_start_in_run = max(idx, rstart) - rstart
                bold_end_in_run = min(phrase_end, rend) - rstart

                if rstart >= phrase_end or rend <= idx:
                    new_runs.append(relem)
                    continue

                # This run overlaps the bold range — split it
                before = rtxt[:bold_start_in_run]
                bold_part = rtxt[bold_start_in_run:bold_end_in_run]
                after = rtxt[bold_end_in_run:]

                for seg_text, is_bold in [
                    (before, False), (bold_part, True), (after, False)
                ]:
                    if not seg_text:
                        continue
                    new_r = copy.deepcopy(relem)
                    t_elems = new_r.findall(f"{{{_W}}}t")
                    if t_elems:
                        t_elems[0].text = seg_text
                        t_elems[0].set(
                            "{http://www.w3.org/XML/1998/namespace}space",
                            "preserve",
                        )
                        for extra in t_elems[1:]:
                            new_r.remove(extra)
                    else:
                        t_elem = etree.SubElement(new_r, f"{{{_W}}}t")
                        t_elem.text = seg_text
                        t_elem.set(
                            "{http://www.w3.org/XML/1998/namespace}space",
                            "preserve",
                        )
                    if is_bold:
                        rpr = new_r.find(f"{{{_W}}}rPr")
                        if rpr is None:
                            rpr = etree.SubElement(new_r, f"{{{_W}}}rPr")
                            new_r.insert(0, rpr)
                        if rpr.find(f"{{{_W}}}b") is None:
                            etree.SubElement(rpr, f"{{{_W}}}b")
                        if rpr.find(f"{{{_W}}}bCs") is None:
                            etree.SubElement(rpr, f"{{{_W}}}bCs")
                    new_runs.append(new_r)

            # Replace runs in paragraph
            if len(new_runs) != len(run_boundaries) or any(
                nr is not orig for nr, (_, _, orig) in zip(new_runs, run_boundaries)
            ):
                for r in runs:
                    elem.remove(r)
                for nr in new_runs:
                    elem.append(nr)
                para_modified = True
                # Rebuild for next phrase in same paragraph
                runs = elem.findall(f"{{{_W}}}r")
                full_text_parts = []
                run_boundaries = []
                pos = 0
                for r in runs:
                    rtxt = "".join(
                        t.text or "" for t in r.findall(f"{{{_W}}}t")
                    )
                    run_boundaries.append((pos, pos + len(rtxt), r))
                    full_text_parts.append(rtxt)
                    pos += len(rtxt)
                full_text = "".join(full_text_parts)

        if para_modified:
            applied += 1

    return applied


def _apply_patches(
    document_xml_bytes: bytes,
    mappings: list[SectionMapping],
    full_md: str = "",
) -> tuple[bytes, int, int, int, int, int, int, int]:
    """Apply Options B, C, C+bold, D, and E to all mapped sections, plus
    table patching on unmatched DOCX sections.

    Returns (modified_xml_bytes, bullets_inserted, paragraphs_corrected,
             table_rows_updated, table_rows_inserted, table_rows_removed,
             bullets_removed, bold_spans_applied).
    """
    root = etree.fromstring(document_xml_bytes)
    body = root.find(f"{{{_W}}}body")
    if body is None:
        return document_xml_bytes, 0, 0, 0, 0

    total_bullets = 0
    total_bullets_removed = 0
    total_corrections = 0
    total_bold_applied = 0
    total_rows_updated = 0
    total_rows_inserted = 0
    total_rows_removed = 0

    for mapping in mappings:
        if mapping.action != "unchanged" or mapping.docx_section is None:
            continue

        heading_text = mapping.docx_section.heading
        try:
            start_idx, end_idx = _find_section_range(body, heading_text)
        except ValueError:
            continue

        # Refresh section_elements after each insertion pass
        section_elements = list(body)[start_idx:end_idx]

        # Option D: table row patching (before Option C to avoid conflicts)
        rows_updated, rows_inserted, rows_removed = _patch_table_rows(
            section_elements, mapping.md_content, body
        )
        if rows_updated or rows_inserted or rows_removed:
            total_rows_updated += rows_updated
            total_rows_inserted += rows_inserted
            total_rows_removed += rows_removed
            parts = []
            if rows_updated:
                parts.append(f"{rows_updated} row(s) updated")
            if rows_inserted:
                parts.append(f"{rows_inserted} row(s) inserted")
            if rows_removed:
                parts.append(f"{rows_removed} stale row(s) removed")
            click.echo(f"  Table patch in '{heading_text}': {', '.join(parts)}")
            # Refresh after table modifications
            start_idx, end_idx = _find_section_range(body, heading_text)
            section_elements = list(body)[start_idx:end_idx]

        corrections = _patch_text_corrections(section_elements, mapping.md_content)
        if corrections:
            total_corrections += corrections
            click.echo(f"  Corrected {corrections} paragraph(s) in '{heading_text}'")

        # Option C+bold: apply **bold** emphasis from source MD
        bold_applied = _patch_bold_emphasis(section_elements, mapping.md_content)
        if bold_applied:
            total_bold_applied += bold_applied
            click.echo(f"  Applied bold to {bold_applied} span(s) in '{heading_text}'")

        bullets = _patch_new_bullets(body, section_elements, end_idx, mapping.md_content)
        if bullets:
            total_bullets += bullets
            click.echo(f"  Inserted {bullets} new bullet(s) in '{heading_text}'")

        # Option E: remove DOCX bullets not present in source MD.
        # Refresh section_elements after Option B insertions.
        try:
            start_idx, end_idx = _find_section_range(body, heading_text)
        except ValueError:
            pass
        else:
            section_elements = list(body)[start_idx:end_idx]
            removed = _remove_stale_bullets(body, section_elements, mapping.md_content)
            if removed:
                total_bullets_removed += removed
                click.echo(f"  Removed {removed} stale bullet(s) in '{heading_text}'")

    # Second pass: table + text corrections for LLM-edited sections.
    # Option D patches stale table cell values the LLM carried over from the
    # original DOCX (e.g. wrong dates).  Option C-edit (0.90 threshold) fixes
    # near-exact paragraph mismatches where the LLM dropped/changed a word.
    for mapping in mappings:
        if mapping.action != "replace" or mapping.docx_section is None:
            continue
        heading_text = mapping.docx_section.heading
        try:
            start_idx, end_idx = _find_section_range(body, heading_text)
        except ValueError:
            continue
        section_elements = list(body)[start_idx:end_idx]

        # Option D on LLM-edited sections: correct stale table cell values
        rows_updated, rows_inserted, rows_removed = _patch_table_rows(
            section_elements, mapping.md_content, body
        )
        if rows_updated or rows_inserted or rows_removed:
            total_rows_updated += rows_updated
            total_rows_inserted += rows_inserted
            total_rows_removed += rows_removed
            parts = []
            if rows_updated:
                parts.append(f"{rows_updated} row(s) updated")
            if rows_inserted:
                parts.append(f"{rows_inserted} row(s) inserted")
            if rows_removed:
                parts.append(f"{rows_removed} stale row(s) removed")
            click.echo(
                f"  Post-LLM table patch in '{heading_text}': {', '.join(parts)}"
            )
            # Refresh after table modifications
            start_idx, end_idx = _find_section_range(body, heading_text)
            section_elements = list(body)[start_idx:end_idx]

        corrections = _patch_text_corrections(
            section_elements, mapping.md_content, min_ratio=0.90
        )
        if corrections:
            total_corrections += corrections
            click.echo(
                f"  Post-LLM corrected {corrections} paragraph(s) in '{heading_text}'"
            )

    # Third pass: Option D on unmatched DOCX sections that contain tables.
    # Some DOCX sections (e.g., sub-headings like "HIGH LEVEL IMPLEMENTATION
    # PLAN") are not matched to any MD section and pass through as-is.  Their
    # tables may contain stale values that the full source MD can correct.
    if full_md:
        mapped_headings = set()
        for m in mappings:
            if m.docx_section is not None:
                mapped_headings.add(m.docx_section.heading)

        # Walk the body to find heading-delimited sections not in the mapping
        children = list(body)
        i = 0
        while i < len(children):
            child = children[i]
            if child.tag != f"{{{_W}}}p":
                i += 1
                continue
            heading_text = _para_heading_text(child)
            if heading_text is None or not heading_text.strip():
                i += 1
                continue
            if heading_text in mapped_headings:
                i += 1
                continue
            # Found an unmatched heading — collect its section elements
            sec_start = i
            sec_end = i + 1
            while sec_end < len(children):
                c = children[sec_end]
                if c.tag == f"{{{_W}}}p" and _para_heading_text(c) is not None:
                    break
                sec_end += 1
            section_elements = children[sec_start:sec_end]
            # Option C: text corrections on unmatched sections
            corrections = _patch_text_corrections(
                section_elements, full_md, min_ratio=0.80
            )
            if corrections:
                total_corrections += corrections
                click.echo(
                    f"  Corrected {corrections} paragraph(s) in "
                    f"'{heading_text}'"
                )
            # Option D: table row patching on unmatched sections
            has_tables = any(
                e.tag == f"{{{_W}}}tbl" for e in section_elements
            )
            if has_tables:
                rows_updated, rows_inserted, rows_removed = _patch_table_rows(
                    section_elements, full_md, body
                )
                if rows_updated or rows_inserted or rows_removed:
                    total_rows_updated += rows_updated
                    total_rows_inserted += rows_inserted
                    total_rows_removed += rows_removed
                    parts = []
                    if rows_updated:
                        parts.append(f"{rows_updated} row(s) updated")
                    if rows_inserted:
                        parts.append(f"{rows_inserted} row(s) inserted")
                    if rows_removed:
                        parts.append(f"{rows_removed} stale row(s) removed")
                    click.echo(
                        f"  Unmatched section table patch in "
                        f"'{heading_text}': {', '.join(parts)}"
                    )
            i = sec_end

    if total_bullets or total_bullets_removed or total_corrections or total_bold_applied or total_rows_updated or total_rows_inserted or total_rows_removed:
        return (
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
            total_bullets,
            total_corrections,
            total_rows_updated,
            total_rows_inserted,
            total_rows_removed,
            total_bullets_removed,
            total_bold_applied,
        )
    return document_xml_bytes, 0, 0, 0, 0, 0, 0, 0


# ---------------------------------------------------------------------------
# Zip repackage
# ---------------------------------------------------------------------------

def _repackage_docx(source: Path, output_path: Path, modified_xml: bytes) -> None:
    """Write a new .docx by copying source entries, replacing document.xml.

    Uses ZipInfo entries directly to preserve per-file metadata (timestamps,
    compression method, entry ordering).

    When source == output_path (--overwrite mode), writes to a temporary file
    first then replaces the original atomically.
    """
    import tempfile

    same_file = source.resolve() == output_path.resolve()
    if same_file:
        fd, tmp_path_str = tempfile.mkstemp(suffix=".docx", dir=output_path.parent)
        os.close(fd)
        dest_path = Path(tmp_path_str)
    else:
        dest_path = output_path

    with zipfile.ZipFile(source, "r") as src_zf:
        with zipfile.ZipFile(dest_path, "w") as dst_zf:
            for item in src_zf.infolist():
                if item.filename == "word/document.xml":
                    dst_zf.writestr(item, modified_xml)
                else:
                    dst_zf.writestr(item, src_zf.read(item.filename))

    if same_file:
        dest_path.replace(output_path)


# ---------------------------------------------------------------------------
# Round-trip validation
# ---------------------------------------------------------------------------

def _count_md_tables(content: str) -> int:
    """Count table blocks in markdown content (separator rows with |---|)."""
    return sum(1 for line in content.split("\n") if re.match(r"\s*\|[\s\-:|]+\|", line))


def _count_md_lines(content: str, prefix: str) -> int:
    """Count lines starting with a specific prefix (e.g. '* ' for bullets)."""
    return sum(1 for line in content.split("\n") if line.strip().startswith(prefix))


def _enforce_bullet_order(
    body: etree._Element,
    mapping: list[SectionMapping],
    edits: list[Edit],
    docx_sections: list,
) -> int:
    """Reorder list paragraphs in LLM-edited sections to match source MD bullet order.

    Returns the number of sections where bullets were reordered.
    """
    from difflib import SequenceMatcher

    edited_headings = {e.target_heading for e in edits if e.kind == "replace"}
    all_section_headings = {s.heading for s in docx_sections}

    # Build a map from DOCX heading → MD bullet texts (in order)
    # Include both replace (LLM-edited) and unchanged sections (Option B
    # may have inserted bullets in wrong positions).
    md_bullets_by_heading: dict[str, list[str]] = {}
    for m in mapping:
        if m.action not in ("replace", "unchanged"):
            continue
        bullets = []
        for bm in _MD_BULLET_RE.finditer(m.md_content):
            cleaned = _clean_md_inline(bm.group(1).strip())
            if cleaned:
                bullets.append(cleaned)
        if len(bullets) >= 2:
            key = m.docx_section.heading if m.docx_section else m.md_heading
            md_bullets_by_heading[key] = bullets

    if not md_bullets_by_heading:
        return 0

    reordered_sections = 0

    # For each section with bullets, find its list paragraphs and reorder
    for heading_text, md_bullets in md_bullets_by_heading.items():
        try:
            start, end = _find_section_range(body, heading_text)
        except ValueError:
            continue

        children = list(body)[start:end]

        # Split list paragraphs into contiguous groups separated by non-list elements.
        # This prevents reordering across sub-section boundaries.
        groups: list[list[tuple[int, etree._Element, str]]] = []
        current_group: list[tuple[int, etree._Element, str]] = []
        for idx_offset, child in enumerate(children):
            if child.tag == f"{{{_W}}}p" and _is_list_para(child):
                txt = "".join(
                    t.text or "" for t in child.findall(f".//{{{_W}}}t")
                ).strip()
                if txt:
                    current_group.append((start + idx_offset, child, txt))
                    continue
            if current_group:
                groups.append(current_group)
                current_group = []
        if current_group:
            groups.append(current_group)

        # Process each contiguous group independently
        for group in groups:
            if len(group) < 2:
                continue

            # Match each list paragraph to the best MD bullet
            matches: list[tuple[int, int]] = []  # (md_idx, group_idx)
            used_md: set[int] = set()
            for gi, (_, _, ptxt) in enumerate(group):
                best_mi = -1
                best_ratio = 0.0
                for mi, mb in enumerate(md_bullets):
                    if mi in used_md:
                        continue
                    ratio = SequenceMatcher(
                        None, ptxt.lower(), mb.lower(), autojunk=False
                    ).ratio()
                    if ratio > best_ratio:
                        best_ratio = ratio
                        best_mi = mi
                if best_mi >= 0 and best_ratio >= 0.60:
                    matches.append((best_mi, gi))
                    used_md.add(best_mi)

            if len(matches) < 2:
                continue

            # Check if already in order
            md_order = [mi for mi, gi in sorted(matches, key=lambda x: x[1])]
            if md_order == sorted(md_order):
                continue

            # Reorder: detach matched list paragraphs, then re-insert in MD order.
            sorted_by_md = sorted(matches, key=lambda x: x[0])
            positions = sorted(group[gi][0] for _, gi in matches)
            desired_elements = [group[gi][1] for _, gi in sorted_by_md]

            # Step 1: Insert placeholder comments at each position
            placeholders = []
            for pos_idx, abs_pos in enumerate(positions):
                elem = list(body)[abs_pos]
                ph = etree.Comment(f"bullet-reorder-{pos_idx}")
                body.insert(list(body).index(elem), ph)
                body.remove(elem)
                placeholders.append(ph)

            # Step 2: Replace placeholders with desired elements in order
            for ph, new_elem in zip(placeholders, desired_elements):
                ph_idx = list(body).index(ph)
                body.insert(ph_idx, new_elem)
                body.remove(ph)

            reordered_sections += 1

    return reordered_sections


def _validate_round_trip(source_md: str, output_path: Path) -> None:
    """Convert output DOCX back to Markdown and compare to source.

    Performs three levels of validation:
    1. Heading structure (count and level distribution)
    2. Section-level content comparison (tables, text similarity, line counts)
    3. Repeated-line preservation (boilerplate that should appear N times)

    Emits warnings for each discrepancy found. Does not block output delivery.
    """
    try:
        from markitdown import MarkItDown
        mid = MarkItDown()
        result = mid.convert(str(output_path))
        output_md = result.text_content
    except Exception as exc:
        click.echo(f"  Info: Round-trip validation skipped ({exc})")
        return

    # Save round-trip .md alongside the output .docx for human review
    md_output_path = output_path.with_suffix(".md")
    try:
        md_output_path.write_text(output_md, encoding="utf-8")
        click.echo(f"  Saved round-trip MD: {md_output_path.name}")
    except Exception as exc:
        click.echo(f"  Info: Could not save round-trip MD ({exc})")

    from ..ai.map import _parse_md_sections

    src_sections = _parse_md_sections(source_md)
    out_sections = _parse_md_sections(output_md)

    click.echo("Validating round-trip output...")

    # --- Level 1: Heading structure ---
    from collections import Counter
    src_levels: Counter = Counter(lvl for _, lvl, _ in src_sections if lvl > 0)
    out_levels: Counter = Counter(lvl for _, lvl, _ in out_sections if lvl > 0)

    src_total = sum(src_levels.values())
    out_total = sum(out_levels.values())

    src_breakdown = ", ".join(f"{v} at H{k}" for k, v in sorted(src_levels.items()))
    out_breakdown = ", ".join(f"{v} at H{k}" for k, v in sorted(out_levels.items()))
    click.echo(f"  Source headings : {src_total} ({src_breakdown})")
    click.echo(f"  Output headings : {out_total} ({out_breakdown})")

    for lvl, src_count in sorted(src_levels.items()):
        out_count = out_levels.get(lvl, 0)
        if out_count < src_count:
            diff = src_count - out_count
            click.echo(
                f"  Warning: {diff} expected H{lvl} section(s) may be missing or "
                f"rendered at wrong level in output"
            )

    if out_total > src_total * 2 and out_total - src_total > 5:
        click.echo(
            f"  Warning: output has {out_total - src_total} more headings than source "
            f"({out_total} vs {src_total}) — LLM may have over-structured content"
        )

    # --- Level 2: Section-level content comparison ---
    # Build lookup: normalized heading -> (heading, content) for output
    out_by_heading: dict[str, tuple[str, str]] = {}
    for h, _lvl, c in out_sections:
        key = _normalize_text(h)
        out_by_heading[key] = (h, c)

    section_warnings: list[str] = []
    for src_h, src_lvl, src_c in src_sections:
        src_key = _normalize_text(src_h)
        if src_key not in out_by_heading:
            continue  # missing section — already covered by heading count check
        _out_h, out_c = out_by_heading[src_key]

        # Table count comparison
        src_tables = _count_md_tables(src_c)
        out_tables = _count_md_tables(out_c)
        if out_tables > src_tables and src_tables == 0:
            section_warnings.append(
                f"'{src_h}': output has {out_tables} table(s) but source has none "
                f"— possible misplaced table"
            )
        elif out_tables > src_tables + 1:
            section_warnings.append(
                f"'{src_h}': output has {out_tables} table(s) vs source {src_tables} "
                f"— possible duplicated table"
            )

        # Text similarity
        from difflib import SequenceMatcher
        src_norm = _normalize_text(src_c)
        out_norm = _normalize_text(out_c)
        if src_norm and out_norm:
            ratio = SequenceMatcher(None, src_norm, out_norm).ratio()
            if ratio < 0.70 and len(src_norm) > 100:
                section_warnings.append(
                    f"'{src_h}': text similarity {ratio:.0%} — "
                    f"significant content change"
                )

        # Content size comparison (detect large additions or losses)
        src_len = len(src_norm) if src_norm else 0
        out_len = len(out_norm) if out_norm else 0
        if src_len > 100:
            if out_len > src_len * 1.5:
                section_warnings.append(
                    f"'{src_h}': output {out_len} chars vs source {src_len} chars "
                    f"({out_len / src_len:.0%}) — possible duplicated content"
                )
            elif out_len < src_len * 0.5:
                section_warnings.append(
                    f"'{src_h}': output {out_len} chars vs source {src_len} chars "
                    f"({out_len / src_len:.0%}) — possible content loss"
                )

        # Bullet count comparison
        src_bullets = _count_md_lines(src_c, "* ")
        out_bullets = _count_md_lines(out_c, "* ")
        if src_bullets > 3 and out_bullets == 0:
            section_warnings.append(
                f"'{src_h}': source has {src_bullets} bullet(s) but output has none"
            )
        elif src_bullets > 5 and out_bullets < src_bullets * 0.5:
            section_warnings.append(
                f"'{src_h}': output has {out_bullets} bullet(s) vs source "
                f"{src_bullets} — possible bullet loss"
            )

    # --- Level 3: Repeated-line preservation ---
    # Lines that appear 3+ times in the source should appear the same number
    # of times in the output (catches boilerplate stripping).
    src_line_counts: Counter = Counter()
    for line in source_md.split("\n"):
        stripped = line.strip()
        if stripped and len(stripped) >= 20:
            src_line_counts[stripped] += 1

    out_line_counts: Counter = Counter()
    for line in output_md.split("\n"):
        stripped = line.strip()
        if stripped and len(stripped) >= 20:
            out_line_counts[stripped] += 1

    for line_text, src_count in src_line_counts.items():
        if src_count >= 3:
            out_count = out_line_counts.get(line_text, 0)
            if out_count < src_count:
                preview = line_text[:60] + ("..." if len(line_text) > 60 else "")
                section_warnings.append(
                    f"repeated line appears {out_count}/{src_count} times: "
                    f"'{preview}'"
                )

    # --- Emit all warnings ---
    if section_warnings:
        click.echo(f"  Section-level issues ({len(section_warnings)}):")
        for w in section_warnings:
            click.echo(f"    - {w}")
    else:
        click.echo("  Info: Section-level content comparison looks consistent.")


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def _fallback_to_create(
    input_path: Path,
    output_path: Path,
    target: Path,
    verbose: bool,
    accept_changes: bool,
    reason: str,
    match_pct: str,
) -> bool:
    """Prompt user and fall back to create mode with base as style reference.

    Returns True if the fallback was executed (caller should return early),
    False if the user declined (caller should continue with update pipeline).
    """
    click.echo(
        f"  Documents appear too different for update mode ({match_pct} match)."
    )
    click.echo(f"  Reason: {reason}")
    if accept_changes:
        click.echo("  Creating new document with base styling (--accept-changes).")
        proceed = True
    else:
        proceed = click.confirm(
            "  Fall back to create mode using base document for styling?",
            default=True,
        )
    if not proceed:
        return False
    click.echo(f"Converting {input_path.name} → {output_path.name} (create mode with style reference)...")
    from .pandoc import run as pandoc_run
    pandoc_run(input_path, output_path, ref_doc=target, toc=False, verbose=verbose)
    return True


def run(
    input_path: Path,
    output_path: Path,
    *,
    target: Path | None,
    accept_changes: bool = False,
    verbose: bool,
) -> None:
    """Apply edits from edit_plan to target .docx and write result to output_path."""
    if target is None:
        raise ValueError("xml_edit approach requires a base .docx file")

    client = get_client_or_none()
    if client is None:
        click.echo("No AI provider configured — running in deterministic-only mode.")
        click.echo("  (Set AI_PROVIDER, AI_MODEL, and API key in .env for AI-augmented updates.)")
    md_text = input_path.read_text(encoding="utf-8")

    # 1. Check for conflicts (deterministic)
    click.echo("Scanning for tracked changes...")
    conflicts = detect_conflicts(target)
    click.echo(f"  {conflicts.summary}")

    # 2. Read and chunk the DOCX
    click.echo("Chunking base document...")
    document_xml = extract_document_xml(target)

    # Build heading and list style maps from styles.xml
    global _heading_map
    styles_xml = extract_styles_xml(target)
    if styles_xml:
        _heading_map = build_heading_style_map(styles_xml)
        list_style_map = build_list_style_map(styles_xml)
    else:
        _heading_map = None
        list_style_map = {"bullet": None, "numbered": None}

    # Warn if list styles could not be detected (LLM will fall back to ListBullet/ListNumber)
    if list_style_map.get("bullet") is None:
        click.echo(
            "  Info: No ListBullet-style paragraph found in styles.xml — "
            "using 'ListBullet' as default. Verify this style exists in the document."
        )
    if list_style_map.get("numbered") is None:
        click.echo(
            "  Info: No ListNumber-style paragraph found in styles.xml — "
            "using 'ListNumber' as default. Verify this style exists in the document."
        )

    if conflicts.has_tracked_changes:
        click.echo("  Accepting all tracked changes...")
        document_xml = _accept_tracked_changes(document_xml)

    docx_sections = chunk_docx_xml(document_xml, heading_map=_heading_map)
    # Filter out empty-heading sections — these are structural artifacts (heading-styled
    # paragraphs with no text) that confuse the mapper and produce empty ## in output.
    empty_count = sum(1 for s in docx_sections if not s.heading.strip())
    if empty_count:
        docx_sections = [s for s in docx_sections if s.heading.strip()]
        click.echo(f"  Filtered {empty_count} empty-heading section(s)")
    click.echo(f"  {len(docx_sections)} sections found")
    if verbose:
        for s in docx_sections:
            click.echo(f"    [H{s.heading_level}] {s.heading!r}")

    # 3. Pre-comparison: convert base DOCX to MD, diff against source MD.
    # This tells us exactly what changed, avoiding unnecessary LLM calls.
    heading_renames: dict[str, str] = {}  # base_heading → source_heading
    precompare_used = False
    try:
        from markitdown import MarkItDown as _MID
        from ..ai.pre_compare import (
            pre_compare, pre_compare_overall_similarity,
            map_sections_precompare,
        )
        click.echo("Pre-comparing base document against source...")
        _mid = _MID()
        _base_result = _mid.convert(str(target))
        _base_md = _base_result.text_content

        _diffs = pre_compare(_base_md, md_text)
        _reliability = pre_compare_overall_similarity(_diffs)

        n_identical = sum(1 for d in _diffs if d.action == "identical")
        n_minor = sum(1 for d in _diffs if d.action == "minor_edit")
        n_major = sum(1 for d in _diffs if d.action == "major_change")
        n_new = sum(1 for d in _diffs if d.action == "new")
        n_renamed = sum(1 for d in _diffs if d.heading_renamed)
        click.echo(
            f"  {len(_diffs)} sections: {n_identical} identical, "
            f"{n_minor} minor edit, {n_major} major change, {n_new} new"
            + (f", {n_renamed} heading rename(s)" if n_renamed else "")
        )

        # Primary divergence check: if very few sections match and most are
        # new, the documents are too different for update mode.
        n_matched = n_identical + n_minor
        if _reliability < 0.25 and n_new > n_matched and len(_diffs) >= 3:
            if _fallback_to_create(
                input_path, output_path, target, verbose, accept_changes,
                reason=f"{n_new} new sections vs {n_matched} matching",
                match_pct=f"{_reliability:.0%}",
            ):
                return

        if _reliability >= 0.50:
            precompare_used = True
            # Collect heading renames for later XML patching.
            for d in _diffs:
                if d.heading_renamed and d.base_heading is not None:
                    heading_renames[d.base_heading] = d.source_heading
            click.echo("Mapping sections (pre-comparison)...")
            mapping = map_sections_precompare(_diffs, docx_sections)
        else:
            click.echo(
                f"  Pre-comparison unreliable ({_reliability:.0%} match) "
                f"— falling back to {'AI' if client else 'deterministic'} mapping"
            )
        # Detect heading renames for non-MD headings: DOCX headings that
        # appear as all-caps body text lines in the source MD.  Only fires
        # when the candidate line shares the DOCX heading's distinctive words
        # (not just boilerplate like "EXHIBIT X TO SOW#").
        if precompare_used:
            from difflib import SequenceMatcher as _SM
            mapped_headings = {
                m.md_heading for m in mapping if m.docx_section is not None
            }
            _used_renames: set[str] = set()
            for ds in docx_sections:
                if ds.heading in mapped_headings:
                    continue
                if ds.heading in heading_renames:
                    continue
                dh_norm = ds.heading.strip().upper()
                if len(dh_norm) < 10:
                    continue
                best_sim = 0.0
                best_line: str | None = None
                for line in md_text.splitlines():
                    sl = line.strip()
                    if not sl or sl.startswith("#") or sl.startswith("|"):
                        continue
                    if sl != sl.upper():
                        continue
                    sl_clean = _clean_md_inline(sl).upper()
                    if sl_clean in _used_renames:
                        continue
                    sim = _SM(None, dh_norm, sl_clean, autojunk=False).ratio()
                    if sim > best_sim:
                        best_sim = sim
                        best_line = sl
                if (
                    best_line
                    and best_sim >= 0.70
                    and best_line.strip().upper() != dh_norm
                    # The candidate must share at least one distinctive word
                    # (3+ chars, not boilerplate like "TO", "SOW#", "THE")
                    and (
                        set(w for w in dh_norm.split() if len(w) >= 4)
                        & set(w for w in _clean_md_inline(best_line).upper().split() if len(w) >= 4)
                    )
                ):
                    heading_renames[ds.heading] = best_line
                    _used_renames.add(_clean_md_inline(best_line).upper())

        del _base_md, _base_result, _mid
    except Exception as exc:
        click.echo(f"  Pre-comparison skipped ({exc})")

    if not precompare_used:
        # Fallback: original AI or deterministic mapping
        if client is not None:
            click.echo("Mapping sections (AI)...")
            mapping = map_sections(client, md_text, docx_sections)
        else:
            click.echo("Mapping sections (deterministic)...")
            mapping = map_sections_deterministic(md_text, docx_sections)

    # Force-preserve synthetic boilerplate sections (e.g. signature pages).
    # These are detected dynamically by both parsers and should never be
    # regenerated by the LLM — the original DOCX content must stay as-is.
    for m in mapping:
        if m.md_heading.startswith("(") and m.md_heading.endswith(")") and m.action != "unchanged":
            m.action = "unchanged"

    # Deterministic unchanged detection: if MD text matches DOCX text after
    # normalization, skip the LLM entirely to preserve original formatting.
    # (When pre-comparison is used, most sections are already "unchanged" —
    # this pass catches any remaining edge cases and is harmless to run twice.)
    det_unchanged = 0
    for m in mapping:
        if m.action == "replace" and _sections_text_match(m):
            m.action = "unchanged"
            det_unchanged += 1
    if det_unchanged:
        click.echo(f"  {det_unchanged} section(s) unchanged (text match) — preserving formatting")

    n_changes = sum(1 for m in mapping if m.action != "unchanged")
    click.echo(f"  {len(mapping)} MD sections mapped, {n_changes} require edits")
    if verbose:
        for m in mapping:
            md_h = m.md_heading.encode("ascii", "replace").decode()
            docx_h = m.docx_section.heading.encode("ascii", "replace").decode() if m.docx_section else ""
            click.echo(
                f"  '{md_h}' -> {m.action}"
                + (f" ('{docx_h}')" if m.docx_section else "")
            )

    # Secondary divergence check (post-mapping): catch cases where
    # pre-compare wasn't available or didn't trigger the primary check.
    # Fires when most sections are inserts and replace sections have low similarity.
    if len(mapping) >= 3:
        n_insert = sum(1 for m in mapping if m.action == "insert")
        insert_ratio = n_insert / len(mapping)
        if insert_ratio > 0.60:
            # Compute average similarity for replace mappings
            replace_sims: list[float] = []
            for m in mapping:
                if m.action == "replace" and m.docx_section is not None:
                    md_n = _normalize_text(m.md_content)
                    dx_n = _normalize_text(
                        _extract_docx_section_text(m.docx_section.xml_fragment)
                    )
                    if md_n and dx_n:
                        from difflib import SequenceMatcher
                        replace_sims.append(
                            SequenceMatcher(None, md_n, dx_n, autojunk=False).ratio()
                        )
            avg_sim = sum(replace_sims) / len(replace_sims) if replace_sims else 0.0
            if avg_sim < 0.30:
                overall_pct = f"{(1 - insert_ratio):.0%}"
                if _fallback_to_create(
                    input_path, output_path, target, verbose, accept_changes,
                    reason=(
                        f"{n_insert}/{len(mapping)} sections unmatched, "
                        f"matched sections avg {avg_sim:.0%} similar"
                    ),
                    match_pct=overall_pct,
                ):
                    return

    # 4. Build edit plan (AI, batched) — skip in deterministic-only mode
    edits: list[Edit] = []
    if client is not None:
        click.echo(f"Building edit plan ({n_changes} section(s) to update)...")
        # Primary: styles actually used in document (frequency-based; most reliable).
        # Fallback: add styles from the full heading map for levels not yet in use — so
        # inserted sections at a level not yet present can still get a proper style.
        # Per design: only use styles that exist in styles.xml; never invent new ones.
        doc_heading_styles_in_use = _extract_heading_styles_in_use(docx_sections)
        if _heading_map:
            covered_levels = set(doc_heading_styles_in_use.values())
            combined_heading_styles: dict[str, int] = dict(doc_heading_styles_in_use)
            for style_id, level in _heading_map.items():
                if level not in covered_levels:
                    combined_heading_styles[style_id] = level
                    covered_levels.add(level)
            doc_heading_styles = combined_heading_styles
        else:
            doc_heading_styles = doc_heading_styles_in_use
        edits = build_edit_plan(
            client, mapping,
            doc_heading_styles=doc_heading_styles,
            doc_list_styles=list_style_map,
        )
        if verbose:
            for e in edits:
                click.echo(f"  {e.kind}: {e.target_heading!r}")
    else:
        # Deterministic mode: check for single-preamble fallback first.
        is_single_preamble = (
            len(docx_sections) == 1
            and docx_sections[0].heading == "(preamble)"
            and len(mapping) <= 1
        )
        if is_single_preamble:
            md_len = len(_normalize_text(md_text))
            docx_text_len = sum(
                len(_normalize_text(_extract_docx_section_text(s.xml_fragment)))
                for s in docx_sections
            )
            is_substantially_different = md_len > docx_text_len * 1.5 or md_len > docx_text_len + 2000
            if is_substantially_different:
                click.echo(
                    "No section structure detected. "
                    "Falling back to create mode with base as style reference."
                )
                from .pandoc import run as pandoc_run
                pandoc_run(
                    input_path, output_path,
                    ref_doc=target, toc=False, verbose=verbose,
                )
                return
            else:
                click.echo("No edits needed — copying base as-is.")
                if target.resolve() != output_path.resolve():
                    shutil.copy2(target, output_path)
                return

        # Force all matched sections to "unchanged" so the patch step (7c)
        # can apply targeted bullet/text/table fixes.
        for m in mapping:
            if m.action == "replace":
                m.action = "unchanged"
        n_matched = sum(1 for m in mapping if m.docx_section is not None)
        click.echo(
            f"  Skipping AI edit plan — {n_matched} section(s) matched deterministically, "
            f"applying targeted patches only."
        )

    if not edits and client is not None and not precompare_used:
        # AI mode produced no edits AND pre-comparison was not used — check for
        # single-preamble fallback or copy as-is.  When pre-comparison drives
        # the pipeline, 0 LLM edits is expected (minor changes handled by
        # deterministic patches) — do NOT bail out.
        md_len = len(_normalize_text(md_text))
        docx_text_len = sum(
            len(_normalize_text(_extract_docx_section_text(s.xml_fragment)))
            for s in docx_sections
        )
        is_single_preamble = (
            len(docx_sections) == 1
            and docx_sections[0].heading == "(preamble)"
            and len(mapping) <= 1
        )
        is_substantially_different = md_len > docx_text_len * 1.5 or md_len > docx_text_len + 2000

        if is_single_preamble and is_substantially_different:
            click.echo(
                "No section structure detected (no ATX headings in MD, no heading styles in DOCX). "
                "Falling back to create mode with base as style reference."
            )
            from .pandoc import run as pandoc_run
            pandoc_run(
                input_path, output_path,
                ref_doc=target, toc=False, verbose=verbose,
            )
            return

        click.echo("No edits needed — copying base as-is.")
        if target.resolve() != output_path.resolve():
            shutil.copy2(target, output_path)
        return

    if edits:
        # 5. Validate media relationship references (deterministic)
        valid_rel_ids = _load_relationship_ids(target)
        _check_media_refs(edits, valid_rel_ids)

        # 6. Summarize and confirm (AI, unless --accept-changes)
        if not accept_changes and client is not None:
            click.echo("Generating change summary...")
            summary = summarize_changes(client, mapping, edits)
            click.echo(f"\nPlanned changes:\n{summary}\n")
            if not click.confirm("Apply these changes?"):
                click.echo("Aborted.")
                sys.exit(0)

    # 7. Apply edits via XML surgery (edits may be empty in deterministic mode)
    click.echo("Applying edits...")
    modified_xml = _apply_edits(document_xml, edits)

    # 7a-pre. Remove LLM-generated headings + trailing content that duplicate
    # an actual preserved section.  The LLM sometimes regenerates the next
    # section's heading and body inside the current edited section.
    # Only fires when: (a) we're inside an edited section, (b) we encounter
    # a heading matching a non-edited section, AND (c) the very next section-
    # boundary heading has the same text (confirming the real section follows
    # immediately and the LLM content is truly a duplicate).
    _root = etree.fromstring(modified_xml)
    _body = _root.find(f"{{{_W}}}body")
    if _body is not None:
        _edited_hdg = {e.target_heading for e in edits if e.kind == "replace"}
        _all_hdg = {s.heading for s in docx_sections}
        _non_edited_hdg = _all_hdg - _edited_hdg
        _in_edited_pre = False
        _dup_remove: list[etree._Element] = []
        _candidate: list[etree._Element] = []
        _candidate_heading: str | None = None
        for _child in list(_body):
            if _child.tag == f"{{{_W}}}p":
                _htxt = _para_heading_text(_child)
                if _htxt is not None and _htxt in _all_hdg:
                    if _candidate_heading is not None:
                        if _htxt == _candidate_heading:
                            # Confirmed: real section follows — remove
                            # the LLM-generated duplicate.
                            _dup_remove.extend(_candidate)
                        _candidate = []
                        _candidate_heading = None
                    if _htxt in _edited_hdg:
                        _in_edited_pre = True
                    elif _in_edited_pre and _htxt in _non_edited_hdg:
                        # Start collecting candidate elements for removal;
                        # will confirm when we see the real heading next.
                        _candidate_heading = _htxt
                        _candidate = [_child]
                    else:
                        _in_edited_pre = False
                    continue
            if _candidate_heading is not None:
                _candidate.append(_child)
        if _dup_remove:
            for _elem in _dup_remove:
                _body.remove(_elem)
            modified_xml = etree.tostring(
                _root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            click.echo(f"  Removed {len(_dup_remove)} LLM-duplicated element(s) from edited section(s)")
            # Re-parse after modification
            _root = etree.fromstring(modified_xml)
            _body = _root.find(f"{{{_W}}}body")

    # 7a. Post-LLM dedup: remove duplicate paragraphs only in LLM-edited
    # sections.
    if _body is not None:
        _dedup_count = 0
        # _original_ids was captured before _apply_edits; elements in the
        # post-edit tree whose id() is NOT in _original_ids are LLM-generated.
        _to_remove: list[etree._Element] = []
        # Walk the body and identify contiguous LLM-edited ranges by
        # using the edit target headings to find section boundaries.
        _edited_headings = {e.target_heading for e in edits if e.kind == "replace"}
        # Use all DOCX section headings as boundaries so sub-headings
        # generated by the LLM inside an edited section don't reset scope.
        _all_section_headings = {s.heading for s in docx_sections}
        _in_edited = False
        _section_seen: set[str] = set()
        for _child in list(_body):
            if _child.tag == f"{{{_W}}}p":
                _htxt = _para_heading_text(_child)
                if _htxt is not None and _htxt in _all_section_headings:
                    _in_edited = _htxt in _edited_headings
                    _section_seen = set()
                    continue
            if not _in_edited:
                continue
            if _child.tag != f"{{{_W}}}p":
                continue
            _text_raw = "".join(
                t.text or "" for t in _child.findall(f".//{{{_W}}}t")
            ).strip()
            if not _text_raw or len(_text_raw) < 40:
                continue
            _text = _normalize_text(_text_raw)
            if _text in _section_seen:
                _to_remove.append(_child)
                _dedup_count += 1
            else:
                _section_seen.add(_text)
        # Second pass: remove paragraphs in edited sections whose text
        # also appears in a non-edited section (LLM leaked next section's
        # content).  Collect non-edited section texts first, then check.
        _preserved_texts: set[str] = set()
        _in_edited2 = False
        for _child in list(_body):
            if _child.tag == f"{{{_W}}}p":
                _htxt = _para_heading_text(_child)
                if _htxt is not None and _htxt in _all_section_headings:
                    _in_edited2 = _htxt in _edited_headings
                    continue
            if _in_edited2 or _child.tag != f"{{{_W}}}p":
                continue
            _text_raw = "".join(
                t.text or "" for t in _child.findall(f".//{{{_W}}}t")
            ).strip()
            if _text_raw and len(_text_raw) >= 40:
                _preserved_texts.add(_normalize_text(_text_raw))
        # Now remove edited-section paragraphs that duplicate preserved text.
        # Also remove orphaned heading-like paragraphs that precede removed
        # duplicates (e.g. LLM-generated "Client Reference" ListNumber para
        # whose body paragraph was the duplicate).
        _in_edited3 = False
        _prev_para: etree._Element | None = None
        for _child in list(_body):
            if _child.tag == f"{{{_W}}}p":
                _htxt = _para_heading_text(_child)
                if _htxt is not None and _htxt in _all_section_headings:
                    _in_edited3 = _htxt in _edited_headings
                    _prev_para = None
                    continue
            if not _in_edited3 or _child.tag != f"{{{_W}}}p":
                _prev_para = None
                continue
            _text_raw = "".join(
                t.text or "" for t in _child.findall(f".//{{{_W}}}t")
            ).strip()
            _text = _normalize_text(_text_raw) if _text_raw else ""
            if _text_raw and len(_text_raw) >= 40 and _text in _preserved_texts:
                # Also remove the preceding paragraph if its text matches
                # a preserved section heading (orphaned LLM heading).
                if _prev_para is not None:
                    _prev_text = "".join(
                        t.text or "" for t in _prev_para.findall(f".//{{{_W}}}t")
                    ).strip()
                    # Remove preceding para if it's an orphaned heading:
                    # either a known section heading or a short non-list para
                    # (likely an LLM-generated sub-heading for the duplicate).
                    _is_orphan_heading = (
                        (_prev_text in _all_section_headings and _prev_text not in _edited_headings)
                        or (0 < len(_prev_text) < 100 and not _is_list_para(_prev_para))
                    )
                    if _is_orphan_heading:
                        _to_remove.append(_prev_para)
                        _dedup_count += 1
                _to_remove.append(_child)
                _dedup_count += 1
                _prev_para = None
            else:
                _prev_para = _child

        # Third pass: remove paragraphs in edited sections whose text matches
        # a preserved section heading (A2 fix — LLM sometimes regenerates
        # headings from adjacent preserved sections like "Client Reference").
        _preserved_headings: set[str] = set()
        for _s in docx_sections:
            if _s.heading not in _edited_headings:
                _preserved_headings.add(_normalize_text(_s.heading))
        if _preserved_headings:
            _in_edited4 = False
            _prev_para4: etree._Element | None = None
            for _child in list(_body):
                if _child.tag == f"{{{_W}}}p":
                    _htxt = _para_heading_text(_child)
                    if _htxt is not None and _htxt in _all_section_headings:
                        _in_edited4 = _htxt in _edited_headings
                        _prev_para4 = None
                        continue
                if not _in_edited4 or _child.tag != f"{{{_W}}}p":
                    _prev_para4 = None
                    continue
                _text_raw4 = "".join(
                    t.text or "" for t in _child.findall(f".//{{{_W}}}t")
                ).strip()
                if not _text_raw4:
                    _prev_para4 = _child
                    continue
                _text4 = _normalize_text(_text_raw4)
                if _text4 in _preserved_headings and _child not in _to_remove:
                    _to_remove.append(_child)
                    _dedup_count += 1
                    # Also remove preceding short non-list paragraph (orphan heading)
                    if _prev_para4 is not None and _prev_para4 not in _to_remove:
                        _pt4 = "".join(
                            t.text or "" for t in _prev_para4.findall(f".//{{{_W}}}t")
                        ).strip()
                        if 0 < len(_pt4) < 100 and not _is_list_para(_prev_para4):
                            _to_remove.append(_prev_para4)
                            _dedup_count += 1
                    _prev_para4 = None
                else:
                    _prev_para4 = _child

        for _elem in _to_remove:
            _body.remove(_elem)
        if _dedup_count:
            modified_xml = etree.tostring(
                _root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            click.echo(f"  Removed {_dedup_count} duplicate paragraph(s)")

    # 7a-rename. Apply heading renames detected by pre-comparison.
    # Updates w:t text in heading paragraphs from old to new heading text.
    if heading_renames:
        _root = etree.fromstring(modified_xml)
        _body = _root.find(f"{{{_W}}}body")
        _renames_applied = 0
        if _body is not None:
            for _child in _body:
                if _child.tag != f"{{{_W}}}p":
                    continue
                _htxt = _para_heading_text(_child)
                if _htxt is None or _htxt not in heading_renames:
                    continue
                _new_heading = heading_renames[_htxt]
                # Replace text in all w:t elements of this paragraph
                _t_elems = _child.findall(f".//{{{_W}}}t")
                if len(_t_elems) == 1:
                    _t_elems[0].text = _new_heading
                elif _t_elems:
                    _t_elems[0].text = _new_heading
                    for _te in _t_elems[1:]:
                        _te.text = ""
                _renames_applied += 1
        if _renames_applied:
            modified_xml = etree.tostring(
                _root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            click.echo(f"  Renamed {_renames_applied} heading(s)")

    # 7b. Post-LLM bullet style injection
    modified_xml, bullets_styled = _inject_bullet_styles(modified_xml, mapping, docx_path=target)
    if bullets_styled:
        click.echo(f"  Injected bullet style on {bullets_styled} paragraph(s)")

    # 7c. Options B, C, D: patch unchanged sections (bullets, text corrections, table rows)
    click.echo("Patching unchanged sections...")
    modified_xml, bullets_added, corrections_made, rows_updated, rows_inserted, rows_removed, bullets_removed, bold_applied = _apply_patches(modified_xml, mapping, full_md=md_text)
    if bullets_added or bullets_removed or corrections_made or bold_applied or rows_updated or rows_inserted or rows_removed:
        parts = [
            f"{bullets_added} bullet(s) added",
            f"{corrections_made} text correction(s)",
            f"{rows_updated} table row(s) updated",
            f"{rows_inserted} table row(s) inserted",
        ]
        if rows_removed:
            parts.append(f"{rows_removed} stale table row(s) removed")
        if bullets_removed:
            parts.append(f"{bullets_removed} stale bullet(s) removed")
        if bold_applied:
            parts.append(f"{bold_applied} bold span(s) applied")
        click.echo(f"  Patches: {', '.join(parts)}")

    # 7d. Post-patch bullet-order enforcement: reorder list paragraphs to
    # match source MD bullet order. Handles both LLM-reordered sections and
    # Option B insertions placed in wrong positions.
    _root = etree.fromstring(modified_xml)
    _body = _root.find(f"{{{_W}}}body")
    if _body is not None:
        _reorder_count = _enforce_bullet_order(_body, mapping, edits, docx_sections)
        if _reorder_count:
            modified_xml = etree.tostring(
                _root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            click.echo(f"  Reordered bullets in {_reorder_count} section(s) to match source")

    # Whole-document XML validation before writing
    _validate_document_xml(modified_xml)

    # 8. Repackage the .docx, preserving zip entry metadata
    _repackage_docx(target, output_path, modified_xml)

    # 9. Round-trip validation: convert output back to MD and compare structure
    _validate_round_trip(md_text, output_path)
