"""
edit_plan.py — Produce a structured list of XML edits from a SectionMapping.

Takes the output of map.py and asks the LLM to generate concrete, XML-safe
edit instructions that xml_edit.py will apply deterministically.

Large documents are processed in batches (BATCH_SIZE sections per LLM call).
Each batch's output is validated fragment-by-fragment; invalid XML triggers a
focused retry with error feedback before falling back to "preserve" (keep original).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Literal

import click
from lxml import etree

from .client import LLMClient, parse_llm_json
from .map import SectionMapping

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_DECL = f'xmlns:w="{_W}"'

BATCH_SIZE = 5              # max sections per edit-plan LLM call
MAX_SECTION_CHARS = 8000    # md_content chars above which a section gets its own batch
EDIT_MAX_TOKENS = 16384
EDIT_MAX_TOKENS_LARGE = 32768  # for sections that get their own batch

_BULLET_RE = re.compile(r"^[ \t]*[-*]\s+", re.MULTILINE)
_NUMBERED_BODY_RE = re.compile(r"^[ \t]*\d+\.\s+", re.MULTILINE)


def _has_bullet_list(content: str) -> bool:
    """Return True if content contains Markdown bullet list items (* or - prefix)."""
    return bool(_BULLET_RE.search(content))


def _has_numbered_list(content: str) -> bool:
    """Return True if content contains Markdown numbered list items (1. prefix in body)."""
    return bool(_NUMBERED_BODY_RE.search(content))


@dataclass
class Edit:
    kind: Literal["insert", "replace", "delete"]
    target_heading: str
    content: str  # OOXML paragraph(s) — valid w:p elements


_SYSTEM = f"""\
You are a Word document XML editor. For each section edit, produce valid OOXML paragraphs.

Input: a JSON object with:
- sections: array of section edits, each with:
  - md_heading: heading text from the Markdown source
  - md_content: full content from the Markdown source (may contain Markdown bullet/list items)
  - action: "replace", "insert", or "delete"
  - target_heading: DOCX heading to act on (or "(end)" for appending new sections)
  - heading_level: integer (1-9) — the heading level this section's title should use
  - has_bullet_list: true if md_content contains bullet list items (* or - prefix)
  - has_numbered_list: true if md_content contains numbered list items (1. prefix in body text)
  - existing_section_summary: structural summary of current DOCX content (styles, paragraph count, etc.)
- document_heading_styles (optional): maps level numbers to OOXML style IDs,
  e.g. {{"level_1": "ArticleHeading", "level_2": "Heading2"}}. Use these styles for heading paragraphs.
- document_list_styles (optional): maps "bullet" and "numbered" to OOXML style IDs,
  e.g. {{"bullet": "ListBullet", "numbered": "ListNumber"}}. Use these styles for list items.

Respond with a JSON array — no markdown fences, no extra text. Each element:
{{
  "target_heading": "<heading text of the DOCX section to act on>",
  "kind": "replace" | "insert" | "delete" | "preserve",
  "xml_content": "<OOXML paragraphs as a string>"
}}

Use "preserve" only if you cannot produce valid OOXML — the original section will be kept unchanged.

OOXML rules:
- Every paragraph must be a valid <w:p {_NS_DECL}> element.
- HEADING STYLES — CRITICAL: Use the style from document_heading_styles matching heading_level
  (e.g. heading_level=1 → level_1 style, heading_level=2 → level_2 style). If the exact level
  is not in document_heading_styles, fall back to Heading{{N}} (e.g. Heading2 for level 2).
  Example: heading_level=1 with level_1="ArticleHeading":
    <w:p {_NS_DECL}><w:pPr><w:pStyle w:val="ArticleHeading"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>
  Example: heading_level=2 with level_2="Heading2":
    <w:p {_NS_DECL}><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>
- LIST FORMATTING — CRITICAL: When has_bullet_list is true, or md_content has lines starting
  with "* " or "- ", render each such item using the bullet style. NEVER render list items as
  plain Normal paragraphs. Use document_list_styles.bullet if provided, else "ListBullet":
    <w:p {_NS_DECL}><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>
  When has_numbered_list is true (body numbered items, not headings), use
  document_list_styles.numbered if provided, else "ListNumber":
    <w:p {_NS_DECL}><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>
- Normal text: <w:p {_NS_DECL}><w:r><w:t>text</w:t></w:r></w:p>
- Bold: <w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r>
- Italic: <w:r><w:rPr><w:i/></w:rPr><w:t>text</w:t></w:r>
- Table: use standard OOXML w:tbl > w:tr > w:tc > w:p structure with {_NS_DECL} on w:tbl.
- For "delete", xml_content must be an empty string.
- For "insert", target_heading MUST be an EXISTING DOCX heading (one that was in the original
  document, not one being inserted by another edit). Use "(end)" to append at document end.
  Never chain inserts by targeting a section that is itself being inserted.
- Match the styles in existing_section_summary["styles_used"] when generating replacement content.
- BLANK PARAGRAPHS — CRITICAL: Blank paragraphs (`<w:p/>`) serve as visual spacing and MUST be
  preserved. `existing_section_summary.blank_paragraph_count` tells you how many blank paragraphs
  the original section had. Include that many `<w:p/>` elements distributed proportionally through
  your replacement XML (e.g. after the heading, between sub-sections, at section end).
  Example blank paragraph: <w:p {_NS_DECL}/>
- CONTENT COMPLETENESS: Include ALL paragraphs and list items from md_content in your output.
  Do not skip or omit any text. Numbered list items in md_content must each become a separate
  paragraph in the output XML (do not merge them).
- Preserve ALL whitespace in text runs using xml:space="preserve": <w:t xml:space="preserve"> </w:t>
- CRITICAL: xml_content must be a SINGLE-LINE string. Do NOT include literal newlines or
  carriage returns inside xml_content. Write compact XML — all elements on one line.
  Use \\n escape sequences if you must represent newlines, but prefer compact XML.
"""


def _summarize_xml_section(xml_fragment: str) -> dict:
    """Return a compact structural summary of a DOCX XML section.

    Sent to the LLM instead of raw XML to reduce token usage while still
    giving the model enough context to match styles in replacement content.
    """
    try:
        wrapped = f'<root xmlns:w="{_W}">{xml_fragment}</root>'
        root = etree.fromstring(wrapped.encode("utf-8"))

        styles: set[str] = set()
        texts: list[str] = []
        has_tables = False
        para_count = 0
        blank_para_count = 0

        for elem in root.iter():
            if callable(elem.tag):
                continue
            local = etree.QName(elem.tag).localname
            if local == "tbl":
                has_tables = True
            elif local == "p":
                para_count += 1
                ppr = elem.find(f"{{{_W}}}pPr")
                if ppr is not None:
                    ps = ppr.find(f"{{{_W}}}pStyle")
                    if ps is not None:
                        val = ps.get(f"{{{_W}}}val", "")
                        if val:
                            styles.add(val)
                para_text = "".join(t.text or "" for t in elem.findall(f".//{{{_W}}}t")).strip()
                if not para_text:
                    blank_para_count += 1
            elif local == "t" and elem.text:
                texts.append(elem.text)

        return {
            "paragraph_count": para_count,
            "blank_paragraph_count": blank_para_count,
            "styles_used": sorted(styles),
            "has_tables": has_tables,
            "text_preview": "".join(texts)[:150],
        }
    except Exception:
        return {"note": "existing content not parseable"}


def _validate_fragment(xml_str: str) -> None:
    """Validate that xml_str is well-formed XML. Raises lxml.etree.XMLSyntaxError if not."""
    if not xml_str:
        return  # empty is valid for delete operations
    wrapped = f'<root xmlns:w="{_W}">{xml_str}</root>'
    etree.fromstring(wrapped.encode("utf-8"))


def _repair_fragment(
    client: LLMClient,
    original_entry: dict,
    bad_xml: str,
    error_msg: str,
) -> str | None:
    """Retry a single section with error feedback at temperature=0.

    Returns corrected XML string, or None if the LLM signals preserve or
    the retry also fails.
    """
    bad_response_json = json.dumps([{
        "target_heading": original_entry["target_heading"],
        "kind": original_entry["action"],
        "xml_content": bad_xml,
    }])
    retry_messages = [
        {"role": "user", "content": json.dumps([original_entry], indent=2)},
        {"role": "assistant", "content": bad_response_json},
        {
            "role": "user",
            "content": (
                f"Your XML is malformed: {error_msg}\n\n"
                "Return a corrected JSON array with exactly one element. "
                'If you cannot produce valid OOXML, set kind to "preserve" with empty xml_content.'
            ),
        },
    ]
    try:
        raw = client.complete(system=_SYSTEM, messages=retry_messages, max_tokens=EDIT_MAX_TOKENS)
        data = parse_llm_json(raw)
        if not data or data[0].get("kind") == "preserve":
            return None
        fragment = data[0].get("xml_content", "")
        _validate_fragment(fragment)
        return fragment
    except Exception:
        return None


def _compute_insert_targets(mapping: list[SectionMapping]) -> dict[str, str]:
    """For each insert section, determine which existing DOCX heading to insert after.

    Walks the full mapping in MD order. Inserts should appear after the last
    DOCX section seen before them in the source order. Consecutive inserts
    after the same anchor all target that same anchor — _apply_edits inserts
    them in reverse edit order so the final document order is correct.
    """
    targets: dict[str, str] = {}
    last_docx_heading: str | None = None
    for m in mapping:
        if m.action in ("replace", "unchanged") and m.docx_section:
            last_docx_heading = m.docx_section.heading
        elif m.action == "insert":
            targets[m.md_heading] = last_docx_heading if last_docx_heading else "(end)"
    return targets


def _process_batch(
    client: LLMClient,
    batch: list[SectionMapping],
    doc_heading_styles: dict[str, int] | None = None,
    doc_list_styles: dict[str, str | None] | None = None,
    insert_targets: dict[str, str] | None = None,
    max_tokens: int = EDIT_MAX_TOKENS,
) -> list[Edit]:
    """Process one batch of SectionMappings into a list of Edits."""
    sections_for_llm = []
    for m in batch:
        if m.docx_section:
            target = m.docx_section.heading
        elif insert_targets:
            target = insert_targets.get(m.md_heading, "(end)")
        else:
            target = "(end)"
        summary = _summarize_xml_section(m.docx_section.xml_fragment) if m.docx_section else None
        # Heading level: use DOCX section level for replace (preserves structure);
        # use MD-detected level for insert (follows source intent).
        eff_level = (
            m.docx_section.heading_level
            if m.docx_section and m.action == "replace"
            else m.md_heading_level
        )
        sections_for_llm.append({
            "md_heading": m.md_heading,
            "md_content": m.md_content,
            "action": m.action,
            "target_heading": target,
            "heading_level": eff_level,
            "has_bullet_list": _has_bullet_list(m.md_content),
            "has_numbered_list": _has_numbered_list(m.md_content),
            "existing_section_summary": summary,
        })

    # Build a level→style map for the LLM (e.g. {1: "ArticleHeading", 2: "Heading2"}).
    # doc_heading_styles contains only styles present in the document (in-use or defined),
    # so the first style seen at each level is the one to use.
    heading_styles_by_level: dict[int, str] = {}
    if doc_heading_styles:
        for style_id, level in doc_heading_styles.items():
            if level not in heading_styles_by_level:
                heading_styles_by_level[level] = style_id

    user_payload: dict = {"sections": sections_for_llm}
    if heading_styles_by_level:
        user_payload["document_heading_styles"] = {
            f"level_{lvl}": style for lvl, style in sorted(heading_styles_by_level.items())
        }
    if doc_list_styles:
        clean_list_styles = {k: v for k, v in doc_list_styles.items() if v is not None}
        if clean_list_styles:
            user_payload["document_list_styles"] = clean_list_styles

    raw = client.complete(
        system=_SYSTEM,
        messages=[{"role": "user", "content": json.dumps(user_payload, indent=2)}],
        max_tokens=max_tokens,
    )

    user_content_str = json.dumps(user_payload, indent=2)

    try:
        edits_data = parse_llm_json(raw)
    except json.JSONDecodeError as exc:
        retry_msgs = [
            {"role": "user", "content": user_content_str},
            {"role": "assistant", "content": raw},
            {
                "role": "user",
                "content": (
                    f"Your response was not valid JSON: {exc}. "
                    "Return only a valid JSON array, no extra text or markdown fences."
                ),
            },
        ]
        try:
            raw2 = client.complete(system=_SYSTEM, messages=retry_msgs, max_tokens=max_tokens)
            edits_data = parse_llm_json(raw2)
        except (json.JSONDecodeError, Exception) as exc2:
            # Unrecoverable — preserve all sections in this batch
            headings = [m.md_heading for m in batch]
            click.echo(
                f"  Warning: edit plan batch failed after retry ({exc2}). "
                f"Preserving original content for: {', '.join(repr(h) for h in headings[:3])}"
                + (f" (+{len(headings)-3} more)" if len(headings) > 3 else "")
            )
            return []

    # Build lookup so we can pass original entry to repair function
    entries_by_target = {e["target_heading"]: e for e in sections_for_llm}

    result: list[Edit] = []
    for ed in edits_data:
        kind = ed.get("kind", "replace")
        target = ed.get("target_heading", "")
        xml_content = ed.get("xml_content", "")

        # "preserve" means the LLM is declining — keep original, emit no edit
        if kind == "preserve":
            continue

        # Validate fragment XML; retry then fall back to preserve on failure
        if kind != "delete" and xml_content:
            try:
                _validate_fragment(xml_content)
            except etree.XMLSyntaxError as e:
                click.echo(f"  Warning: invalid XML for '{target}', retrying...")
                original_entry = entries_by_target.get(target)
                if original_entry:
                    xml_content = _repair_fragment(client, original_entry, xml_content, str(e))
                    if xml_content is None:
                        click.echo(f"  Warning: retry failed for '{target}', preserving original.")
                        continue
                else:
                    continue

        # Warn on bullet loss: count MD bullets vs ListBullet paragraphs in generated XML
        if kind != "delete" and xml_content:
            entry = entries_by_target.get(target)
            if entry and entry.get("has_bullet_list"):
                md_bullets = len(_BULLET_RE.findall(entry["md_content"]))
                xml_bullets = xml_content.count("ListBullet")
                if md_bullets > 0 and xml_bullets < md_bullets * 0.5:
                    click.echo(
                        f"  Warning: bullet loss in '{target}': "
                        f"{xml_bullets} ListBullet paragraph(s) for {md_bullets} MD bullet(s)."
                    )

        result.append(Edit(kind=kind, target_heading=target, content=xml_content or ""))

    return result


def _make_batches(edits_needed: list[SectionMapping]) -> list[list[SectionMapping]]:
    """Group sections into batches by content size.

    Sections with md_content > MAX_SECTION_CHARS are placed alone (they need
    the full token budget for OOXML generation). Smaller sections are grouped
    up to BATCH_SIZE per batch.
    """
    batches: list[list[SectionMapping]] = []
    current: list[SectionMapping] = []

    for m in edits_needed:
        if len(m.md_content) > MAX_SECTION_CHARS:
            # Flush current batch first, then large section gets its own
            if current:
                batches.append(current)
                current = []
            batches.append([m])
        else:
            current.append(m)
            if len(current) >= BATCH_SIZE:
                batches.append(current)
                current = []

    if current:
        batches.append(current)

    return batches


MAX_PARALLEL_BATCHES = 4  # concurrent LLM calls for edit plan


def build_edit_plan(
    client: LLMClient,
    mapping: list[SectionMapping],
    doc_heading_styles: dict[str, int] | None = None,
    doc_list_styles: dict[str, str | None] | None = None,
) -> list[Edit]:
    """Convert a SectionMapping list into a concrete list of Edits."""
    edits_needed = [m for m in mapping if m.action != "unchanged"]
    if not edits_needed:
        return []

    insert_targets = _compute_insert_targets(mapping)
    batches = _make_batches(edits_needed)

    if len(batches) <= 1:
        # Single batch — no parallelism needed
        if batches:
            is_large = len(batches[0]) == 1 and len(batches[0][0].md_content) > MAX_SECTION_CHARS
            return _process_batch(
                client, batches[0],
                doc_heading_styles=doc_heading_styles,
                doc_list_styles=doc_list_styles,
                insert_targets=insert_targets,
                max_tokens=EDIT_MAX_TOKENS_LARGE if is_large else EDIT_MAX_TOKENS,
            )
        return []

    # Multiple batches — process in parallel
    import concurrent.futures

    click.echo(f"  Processing {len(batches)} batches (up to {MAX_PARALLEL_BATCHES} in parallel)...")

    def _run_batch(batch_idx: int, batch: list[SectionMapping]) -> tuple[int, list[Edit]]:
        is_large = len(batch) == 1 and len(batch[0].md_content) > MAX_SECTION_CHARS
        large = " (large)" if is_large else ""
        click.echo(f"  Edit plan: batch {batch_idx + 1}/{len(batches)}{large}")
        edits = _process_batch(
            client, batch,
            doc_heading_styles=doc_heading_styles,
            doc_list_styles=doc_list_styles,
            insert_targets=insert_targets,
            max_tokens=EDIT_MAX_TOKENS_LARGE if is_large else EDIT_MAX_TOKENS,
        )
        return batch_idx, edits

    results: dict[int, list[Edit]] = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_PARALLEL_BATCHES) as pool:
        futures = {
            pool.submit(_run_batch, idx, batch): idx
            for idx, batch in enumerate(batches)
        }
        for future in concurrent.futures.as_completed(futures):
            idx, edits = future.result()
            results[idx] = edits

    # Reassemble in original batch order
    all_edits: list[Edit] = []
    for idx in range(len(batches)):
        all_edits.extend(results.get(idx, []))

    return all_edits
