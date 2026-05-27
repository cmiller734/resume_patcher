#!/usr/bin/env python3
"""
resume_patcher.py

Deterministic DOCX resume patcher:
- preserve source DOCX formatting/layout as the source of truth
- patch text from external replacements JSON
- avoid rebuilding DOCX from scratch
"""
from __future__ import annotations

import argparse
import json
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph


BULLET_PREFIX_RE = re.compile(r"^\s*(?:\u2022|-|\*)\s+")
DATE_TOKEN_RE = re.compile(
    r"\b(?:jan\.?|january|feb\.?|february|mar\.?|march|apr\.?|april|may|jun\.?|june|"
    r"jul\.?|july|aug\.?|august|sep\.?|sept\.?|september|oct\.?|october|nov\.?|"
    r"november|dec\.?|december|present|\d{4})\b",
    re.IGNORECASE,
)
LOCATION_TOKEN_RE = re.compile(
    r"\b(?:remote|hybrid|onsite|on-site|[A-Z][a-z]+,\s*[A-Z]{2}|[A-Z]{2}\s*/\s*remote)\b",
    re.IGNORECASE,
)

SUPPORTED_STYLE_SOURCES = {
    "section_heading",
    "role_heading",
    "company_heading",
    "date_location",
    "normal",
    "bullet",
    "keep",
}
NON_SECTION_STYLE_SOURCES = {"normal", "role_heading", "company_heading", "date_location"}
ROLE_KEYWORDS = (
    "engineer",
    "developer",
    "architect",
    "consultant",
    "support",
    "services",
    "founder",
    "lead",
    "manager",
    "intern",
)


@dataclass(frozen=True)
class PatchPara:
    text: str
    style_source: str = "normal"


def warn(message: str) -> None:
    print(f"WARNING: {message}")


def apply_resume_bullet_indent(paragraph: Paragraph) -> None:
    """Apply bullet indentation to a paragraph.

    This helper is retained for compatibility, but the patcher primarily preserves
    bullet formatting by cloning existing source DOCX bullet paragraphs.
    """
    paragraph.paragraph_format.left_indent = Inches(0.5)
    paragraph.paragraph_format.first_line_indent = Inches(-0.25)

    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)

    ind.set(qn("w:left"), "720")
    ind.set(qn("w:hanging"), "360")


def _copy_run_format(src_run, dst_run) -> None:
    """Copy formatting from src_run to dst_run without copying text."""
    if src_run is None:
        return
    if src_run._r.rPr is not None:
        dst_run._r.insert(0, deepcopy(src_run._r.rPr))


def _template_runs_by_line(template: Paragraph | None) -> list:
    """Return the first visible run on each visual line of a template paragraph."""
    if template is None:
        return []

    line_runs = [None]
    line_idx = 0

    for run in template.runs:
        for child in run._r.iterchildren():
            if child.tag == qn("w:t"):
                value = child.text or ""
                if value.replace("\n", "").strip() and line_runs[line_idx] is None:
                    line_runs[line_idx] = run
            elif child.tag in {qn("w:br"), qn("w:cr")}:
                line_idx += 1
                while len(line_runs) <= line_idx:
                    line_runs.append(None)

        extra_breaks = run.text.count("\n")
        has_xml_break = any(child.tag in {qn("w:br"), qn("w:cr")} for child in run._r.iterchildren())
        if extra_breaks and not has_xml_break:
            for _ in range(extra_breaks):
                line_idx += 1
                while len(line_runs) <= line_idx:
                    line_runs.append(None)

    fallback = next((r for r in line_runs if r is not None), None)
    filled = []
    previous = fallback
    for r in line_runs:
        if r is not None:
            previous = r
            filled.append(r)
        else:
            filled.append(previous)
    return filled


def set_paragraph_text_keep_format(paragraph: Paragraph, text: str, template: Paragraph | None = None) -> None:
    """Replace paragraph text while preserving paragraph style and run formatting."""
    template = template or paragraph
    line_format_runs = _template_runs_by_line(template)
    fallback_run = (
        template.runs[0]
        if template is not None and template.runs
        else (paragraph.runs[0] if paragraph.runs else None)
    )

    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)

    parts = text.split("\n")
    previous_run = None
    for i, part in enumerate(parts):
        if i > 0 and previous_run is not None:
            previous_run.add_break()

        run = paragraph.add_run(part)
        src_run = line_format_runs[i] if i < len(line_format_runs) and line_format_runs[i] is not None else fallback_run
        _copy_run_format(src_run, run)
        previous_run = run


def _template_line_run_segments(template: Paragraph) -> list[list[tuple[str, Any]]]:
    lines: list[list[tuple[str, Any]]] = [[]]
    for run in template.runs:
        parts = run.text.split("\n")
        for idx, part in enumerate(parts):
            if idx > 0:
                lines.append([])
            if part:
                lines[-1].append((part, run))
    return lines


def _skill_line_format_runs(template: Paragraph, line_idx: int) -> tuple[Any, Any]:
    lines = _template_line_run_segments(template)
    if not lines:
        return None, None

    segments = lines[min(line_idx, len(lines) - 1)]
    if not segments:
        fallback = template.runs[0] if template.runs else None
        return fallback, fallback

    combined_text = "".join(text for text, _run in segments)
    colon_idx = combined_text.find(":")
    label_run = next((run for text, run in segments if text.strip()), segments[0][1])
    rest_run = None

    if colon_idx >= 0:
        pos = 0
        for text, run in segments:
            next_pos = pos + len(text)
            if pos <= colon_idx < next_pos:
                label_run = run
            if next_pos > colon_idx + 1 and text[max(0, colon_idx + 1 - pos) :].strip():
                rest_run = run
                break
            pos = next_pos

    if rest_run is None:
        rest_run = next((run for text, run in segments if text.strip() and run is not label_run), label_run)

    return label_run, rest_run


def set_skills_text_keep_label_format(paragraph: Paragraph, text: str, template: Paragraph) -> None:
    """Replace skills text while preserving the template's label/rest run pattern."""
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)

    previous_run = None
    for line_idx, line in enumerate(text.split("\n")):
        if line_idx > 0:
            if previous_run is None:
                previous_run = paragraph.add_run("")
            previous_run.add_break()

        label_source, rest_source = _skill_line_format_runs(template, line_idx)
        colon_idx = line.find(":")
        if colon_idx > 0:
            label_run = paragraph.add_run(line[: colon_idx + 1])
            _copy_run_format(label_source, label_run)
            previous_run = label_run

            rest = line[colon_idx + 1 :]
            if rest:
                rest_run = paragraph.add_run(rest)
                _copy_run_format(rest_source, rest_run)
                previous_run = rest_run
        else:
            run = paragraph.add_run(line)
            _copy_run_format(rest_source or label_source, run)
            previous_run = run


def insert_paragraph_after(paragraph: Paragraph, text: str, template: Paragraph) -> Paragraph:
    """Insert a paragraph after another paragraph, copying style and paragraph properties."""
    new_p = deepcopy(template._p)
    paragraph._p.addnext(new_p)
    inserted = Paragraph(new_p, paragraph._parent)
    inserted.style = template.style
    set_paragraph_text_keep_format(inserted, text, template)
    return inserted


def delete_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def find_para_index(doc: Document, exact_text: str) -> int:
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == exact_text.strip():
            return i
    raise ValueError(f"Could not find paragraph with exact text: {exact_text!r}")


def find_para_contains(doc: Document, needle: str, start: int = 0, *, case_sensitive: bool = False) -> int:
    target = needle if case_sensitive else needle.lower()
    for i, p in enumerate(doc.paragraphs[start:], start=start):
        haystack = p.text if case_sensitive else p.text.lower()
        if target in haystack:
            return i
    raise ValueError(f"Could not find paragraph containing: {needle!r}")


def replace_block(
    doc: Document,
    start_idx: int,
    end_idx_exclusive: int,
    items: Sequence[PatchPara],
    *,
    style_templates: dict[str, Paragraph],
) -> None:
    """Replace a paragraph block [start_idx, end_idx_exclusive) with items."""
    if not items:
        raise ValueError("replace_block requires at least one item")

    first = doc.paragraphs[start_idx]
    end_anchor = doc.paragraphs[end_idx_exclusive]._p

    template_xml_by_style = {
        style_source: deepcopy(template._p)
        for style_source, template in style_templates.items()
    }
    template_style_name_by_style = {
        style_source: paragraph_style_name(template)
        for style_source, template in style_templates.items()
    }

    def make_template(style_source: str, current: Paragraph | None = None) -> Paragraph:
        if style_source == "keep":
            return current or Paragraph(deepcopy(template_xml_by_style["normal"]), first._parent)

        def safe_normal_template() -> Paragraph:
            fallback = Paragraph(deepcopy(template_xml_by_style["normal"]), first._parent)
            if paragraph_is_heading_1(fallback):
                warn("safe 'normal' fallback resolved to Heading 1; forcing Word Normal style")
                fallback.style = "Normal"
            return fallback

        if style_source not in template_xml_by_style:
            warn(f"style {style_source!r} is not available; using 'normal'")
            style_source = "normal"

        template = Paragraph(deepcopy(template_xml_by_style[style_source]), first._parent)
        if style_source in NON_SECTION_STYLE_SOURCES and paragraph_is_heading_1(template):
            warn(
                f"refusing to apply Heading 1 template to {style_source!r}; "
                "using safe 'normal' template"
            )
            template = safe_normal_template()
        elif style_source in NON_SECTION_STYLE_SOURCES and "heading 1" in template_style_name_by_style.get(style_source, "").lower():
            warn(
                f"refusing to apply Heading 1 template to {style_source!r}; "
                "using safe 'normal' template"
            )
            template = safe_normal_template()
        return template

    first.style = make_template(items[0].style_source, first).style
    set_paragraph_text_keep_format(first, items[0].text, make_template(items[0].style_source, first))

    while first._p.getnext() is not None and first._p.getnext() is not end_anchor:
        delete_paragraph(Paragraph(first._p.getnext(), first._parent))

    cursor = first
    for item in items[1:]:
        cursor = insert_paragraph_after(cursor, item.text, make_template(item.style_source))


def paragraph_is_bullet(paragraph: Paragraph) -> bool:
    pPr = paragraph._p.pPr
    if pPr is not None and pPr.numPr is not None:
        return True
    style_name = (paragraph.style.name if paragraph.style is not None else "").lower()
    if "bullet" in style_name:
        return True
    return bool(BULLET_PREFIX_RE.match(paragraph.text))


def paragraph_style_name(paragraph: Paragraph) -> str:
    return paragraph.style.name if paragraph.style is not None else ""


def paragraph_style_id(paragraph: Paragraph) -> str:
    return paragraph.style.style_id if paragraph.style is not None else ""


def paragraph_has_heading_level(paragraph: Paragraph, level: int) -> bool:
    style_name = paragraph_style_name(paragraph).lower().replace(" ", "")
    style_id = paragraph_style_id(paragraph).lower().replace(" ", "")
    return style_name == f"heading{level}" or style_id == f"heading{level}"


def paragraph_is_heading_1(paragraph: Paragraph) -> bool:
    return paragraph_has_heading_level(paragraph, 1)


def paragraph_is_any_heading(paragraph: Paragraph) -> bool:
    style_name = paragraph_style_name(paragraph).lower()
    style_id = paragraph_style_id(paragraph).lower()
    return "heading" in style_name or style_id.startswith("heading")


def paragraph_is_document_title_style(paragraph: Paragraph) -> bool:
    style_name = paragraph_style_name(paragraph).lower()
    style_id = paragraph_style_id(paragraph).lower()
    return style_name in {"title", "subtitle"} or style_id in {"title", "subtitle"}


def paragraph_looks_like_section_heading(paragraph: Paragraph) -> bool:
    text = paragraph.text.strip()
    if not text:
        return False

    if paragraph_is_any_heading(paragraph) and paragraph_has_heading_level(paragraph, 1):
        return True

    has_letters = any(ch.isalpha() for ch in text)
    return has_letters and text.upper() == text and len(text) <= 80


def paragraph_looks_like_date_location(paragraph: Paragraph) -> bool:
    text = paragraph.text.strip()
    if not text or paragraph_is_bullet(paragraph) or paragraph_looks_like_section_heading(paragraph):
        return False
    if len(text) > 90:
        return False
    if any(word in text.lower() for word in ROLE_KEYWORDS):
        return False
    has_date = DATE_TOKEN_RE.search(text) is not None
    has_location = LOCATION_TOKEN_RE.search(text) is not None
    return has_date and (has_location or "|" in text or " - " in text or " – " in text)


def text_looks_like_date_location(text: str) -> bool:
    text = text.strip()
    if not text or len(text) > 120:
        return False
    has_date = DATE_TOKEN_RE.search(text) is not None
    has_location = LOCATION_TOKEN_RE.search(text) is not None
    return has_date and (has_location or "|" in text or " - " in text or " – " in text)


def paragraph_looks_like_company_heading(paragraph: Paragraph) -> bool:
    text = paragraph.text.strip()
    if not text or paragraph_is_bullet(paragraph) or paragraph_looks_like_section_heading(paragraph):
        return False
    if paragraph_has_heading_level(paragraph, 2):
        return True
    if DATE_TOKEN_RE.search(text):
        return False
    return " | " in text or " – " in text or " - " in text


def paragraph_looks_like_role_heading(paragraph: Paragraph) -> bool:
    text = paragraph.text.strip()
    if not text or paragraph_is_bullet(paragraph) or paragraph_looks_like_section_heading(paragraph):
        return False
    if paragraph_has_heading_level(paragraph, 3):
        return True
    lowered = text.lower()
    return any(word in lowered for word in ROLE_KEYWORDS) and DATE_TOKEN_RE.search(text) is not None


def paragraph_looks_like_combined_role_date_template(paragraph: Paragraph) -> bool:
    lines = [line.strip() for line in paragraph.text.splitlines() if line.strip()]
    return len(lines) >= 2 and "|" not in lines[0] and text_looks_like_date_location(lines[1])


def first_non_empty_paragraph(doc: Document) -> Paragraph:
    for p in doc.paragraphs:
        if p.text.strip():
            return p
    return doc.paragraphs[0]


def first_matching_paragraph(doc: Document, predicate) -> Paragraph | None:
    for p in doc.paragraphs:
        if p.text.strip() and predicate(p):
            return p
    return None


def derive_line_template(template: Paragraph, line_idx: int) -> Paragraph:
    """Create a one-line template using a specific visual line from a source paragraph."""
    derived = Paragraph(deepcopy(template._p), template._parent)
    lines = template.text.splitlines() or [template.text]
    line_text = lines[line_idx] if line_idx < len(lines) else (lines[-1] if lines else "")
    line_runs = _template_runs_by_line(template)
    source_run = line_runs[line_idx] if line_idx < len(line_runs) else None
    if source_run is None and line_runs:
        source_run = line_runs[-1]

    for run in list(derived.runs):
        run._element.getparent().remove(run._element)

    run = derived.add_run(line_text)
    _copy_run_format(source_run, run)
    return derived


def first_safe_non_heading_paragraph(doc: Document) -> Paragraph:
    for p in doc.paragraphs:
        if (
            p.text.strip()
            and not paragraph_is_heading_1(p)
            and not paragraph_is_bullet(p)
            and not paragraph_is_document_title_style(p)
        ):
            return p
    return first_non_empty_paragraph(doc)


def paragraph_is_safe_normal_template(paragraph: Paragraph) -> bool:
    return (
        bool(paragraph.text.strip())
        and not paragraph_is_any_heading(paragraph)
        and not paragraph_is_document_title_style(paragraph)
        and not paragraph_is_bullet(paragraph)
        and not paragraph_looks_like_date_location(paragraph)
        and not paragraph_looks_like_company_heading(paragraph)
        and not paragraph_looks_like_role_heading(paragraph)
    )


def discover_style_templates(doc: Document) -> dict[str, Paragraph]:
    """Map semantic replacement styles to real paragraph templates in the source DOCX."""
    safe_non_heading = first_safe_non_heading_paragraph(doc)
    combined_role_date_template = first_matching_paragraph(doc, paragraph_looks_like_combined_role_date_template)

    templates: dict[str, Paragraph] = {}
    templates["section_heading"] = (
        first_matching_paragraph(doc, paragraph_looks_like_section_heading)
        or first_non_empty_paragraph(doc)
    )
    templates["normal"] = (
        first_matching_paragraph(doc, paragraph_is_safe_normal_template)
        or safe_non_heading
    )
    templates["bullet"] = (
        first_matching_paragraph(doc, paragraph_is_bullet)
        or templates["normal"]
    )
    templates["company_heading"] = (
        first_matching_paragraph(doc, lambda p: paragraph_has_heading_level(p, 2))
        or first_matching_paragraph(doc, paragraph_looks_like_company_heading)
        or templates["normal"]
    )
    templates["role_heading"] = (
        (derive_line_template(combined_role_date_template, 0) if combined_role_date_template is not None else None)
        or first_matching_paragraph(doc, lambda p: paragraph_has_heading_level(p, 3))
        or first_matching_paragraph(doc, paragraph_looks_like_role_heading)
        or templates["company_heading"]
        or templates["normal"]
    )
    templates["date_location"] = (
        (derive_line_template(combined_role_date_template, 1) if combined_role_date_template is not None else None)
        or first_matching_paragraph(doc, paragraph_looks_like_date_location)
        or templates["normal"]
    )

    for style_source in NON_SECTION_STYLE_SOURCES:
        template = templates[style_source]
        if paragraph_is_heading_1(template):
            warn(
                f"semantic style {style_source!r} resolved to Heading 1; "
                "using a non-section fallback instead"
            )
            templates[style_source] = safe_non_heading

    return templates


def infer_end_index(doc: Document, start_idx: int) -> int | None:
    start_para = doc.paragraphs[start_idx]
    start_style_id = start_para.style.style_id if start_para.style is not None else ""

    for i in range(start_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        if not p.text.strip():
            continue

        if paragraph_looks_like_section_heading(p):
            return i

        style_id = p.style.style_id if p.style is not None else ""
        if style_id == start_style_id and not paragraph_is_bullet(p):
            return i

    return None


def parse_patch_paragraphs(raw_items: Any, context: str) -> list[PatchPara]:
    if not isinstance(raw_items, list):
        raise ValueError(f"{context} must be a list")

    parsed: list[PatchPara] = []
    for idx, raw in enumerate(raw_items):
        if isinstance(raw, str):
            style_source = "normal"
            text = raw
            if BULLET_PREFIX_RE.match(text):
                style_source = "bullet"
                text = BULLET_PREFIX_RE.sub("", text, count=1)
            parsed.append(PatchPara(text=text, style_source=style_source))
            continue

        if isinstance(raw, dict):
            text = raw.get("text")
            if not isinstance(text, str):
                raise ValueError(f"{context}[{idx}].text must be a string")

            style_source = raw.get("style", "normal")
            if style_source not in SUPPORTED_STYLE_SOURCES:
                warn(f"{context}[{idx}] has unknown style {style_source!r}; using 'normal'")
                style_source = "normal"

            if style_source == "bullet":
                text = BULLET_PREFIX_RE.sub("", text, count=1)

            parsed.append(PatchPara(text=text, style_source=style_source))
            continue

        raise ValueError(f"{context}[{idx}] must be either a string or an object")

    return parsed


def apply_paragraph_replacements(doc: Document, replacements: Any) -> int:
    if not replacements:
        return 0
    if not isinstance(replacements, list):
        raise ValueError("paragraph_replacements must be a list")

    applied = 0
    for idx, item in enumerate(replacements):
        if not isinstance(item, dict):
            warn(f"paragraph_replacements[{idx}] is not an object; skipping")
            continue

        find = item.get("find")
        replace = item.get("replace")
        if not isinstance(find, str) or not isinstance(replace, str):
            warn(f"paragraph_replacements[{idx}] requires string 'find' and 'replace'; skipping")
            continue

        try:
            para_idx = find_para_index(doc, find)
        except ValueError:
            warn(f"paragraph_replacements[{idx}] could not find exact text: {find!r}")
            continue

        target = doc.paragraphs[para_idx]
        set_paragraph_text_keep_format(target, replace, target)
        applied += 1

    return applied


def find_summary_paragraph_index(doc: Document, heading_text: str = "SUMMARY") -> int | None:
    try:
        heading_idx = find_para_contains(doc, heading_text)
    except ValueError:
        return None

    for i in range(heading_idx + 1, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip():
            return i
    return None


def apply_summary_replacement(doc: Document, summary: Any) -> int:
    if summary is None:
        return 0

    if isinstance(summary, str):
        target_idx = find_summary_paragraph_index(doc)
        if target_idx is None:
            warn("summary was provided but no SUMMARY heading/paragraph target was found")
            return 0

        target = doc.paragraphs[target_idx]
        set_paragraph_text_keep_format(target, summary, target)
        return 1

    if isinstance(summary, dict):
        text = summary.get("text") or summary.get("replace")
        if not isinstance(text, str):
            warn("summary object must include string 'text' (or 'replace'); skipping")
            return 0

        find_text = summary.get("find")
        if isinstance(find_text, str):
            try:
                para_idx = find_para_index(doc, find_text)
            except ValueError:
                warn(f"summary.find exact text not found: {find_text!r}")
                return 0
            target = doc.paragraphs[para_idx]
            set_paragraph_text_keep_format(target, text, target)
            return 1

        heading = summary.get("heading", "SUMMARY")
        if not isinstance(heading, str):
            warn("summary.heading must be a string when provided; using 'SUMMARY'")
            heading = "SUMMARY"
        target_idx = find_summary_paragraph_index(doc, heading)
        if target_idx is None:
            warn(f"summary heading not found: {heading!r}")
            return 0

        target = doc.paragraphs[target_idx]
        set_paragraph_text_keep_format(target, text, target)
        return 1

    warn("summary must be a string or object; skipping")
    return 0


def apply_skills_replacements(doc: Document, skills: Any) -> int:
    if skills is None:
        return 0

    applied = 0

    if isinstance(skills, dict):
        for label, value in skills.items():
            if not isinstance(label, str) or not isinstance(value, str):
                warn(f"skills entry {label!r} is not string:string; skipping")
                continue

            target_prefix = f"{label.strip()}:"
            target_idx = None
            for i, p in enumerate(doc.paragraphs):
                if p.text.strip().lower().startswith(target_prefix.lower()):
                    target_idx = i
                    break

            if target_idx is None:
                warn(f"skills label not found in DOCX: {label!r}")
                continue

            replacement_text = value if value.strip().lower().startswith(target_prefix.lower()) else f"{target_prefix} {value}"
            target = doc.paragraphs[target_idx]
            template = Paragraph(deepcopy(target._p), target._parent)
            set_skills_text_keep_label_format(target, replacement_text, template)
            applied += 1

        return applied

    if isinstance(skills, list):
        try:
            heading_idx = find_para_contains(doc, "SKILLS")
        except ValueError:
            warn("skills list provided but SKILLS heading not found")
            return 0

        target_indices: list[int] = []
        i = heading_idx + 1
        while i < len(doc.paragraphs) and len(target_indices) < len(skills):
            if doc.paragraphs[i].text.strip():
                target_indices.append(i)
            i += 1

        for n, value in enumerate(skills):
            if not isinstance(value, str):
                warn(f"skills[{n}] is not a string; skipping")
                continue
            if n >= len(target_indices):
                warn("skills list has more entries than detected skill target rows; extra entries skipped")
                break
            target = doc.paragraphs[target_indices[n]]
            template = Paragraph(deepcopy(target._p), target._parent)
            set_skills_text_keep_label_format(target, value, template)
            applied += 1

        return applied

    warn("skills must be an object or list; skipping")
    return 0


def apply_block_items(
    doc: Document,
    start_idx: int,
    end_idx: int,
    items: list[PatchPara],
    context: str,
    style_templates: dict[str, Paragraph],
) -> bool:
    if not items:
        warn(f"{context} had no replacement_paragraphs after parsing; skipping")
        return False

    if end_idx <= start_idx:
        warn(f"{context} has invalid range start={start_idx}, end={end_idx}; skipping")
        return False

    end_anchor_text = doc.paragraphs[end_idx].text.strip()
    if items and items[-1].text.strip() == end_anchor_text:
        items = items[:-1]
        if not items:
            warn(f"{context} only repeated the end heading anchor; skipping")
            return False

    replace_block(
        doc,
        start_idx,
        end_idx,
        items,
        style_templates=style_templates,
    )
    return True


def apply_named_blocks(
    doc: Document,
    raw_blocks: Any,
    *,
    label: str,
    style_templates: dict[str, Paragraph],
    default_use_heading_key: bool = False,
    require_end_heading: bool = False,
) -> int:
    if not raw_blocks:
        return 0
    if not isinstance(raw_blocks, list):
        raise ValueError(f"{label} must be a list")

    applied = 0
    for idx, block in enumerate(raw_blocks):
        context = f"{label}[{idx}]"
        if not isinstance(block, dict):
            warn(f"{context} is not an object; skipping")
            continue

        if default_use_heading_key:
            start_heading = block.get("start_heading") or block.get("heading")
        else:
            start_heading = block.get("start_heading")

        if not isinstance(start_heading, str) or not start_heading.strip():
            warn(f"{context} requires a non-empty 'start_heading' or 'heading'; skipping")
            continue

        replacement_paragraphs = block.get("replacement_paragraphs")
        if replacement_paragraphs is None:
            replacement_paragraphs = block.get("content")
        if replacement_paragraphs is None:
            warn(f"{context} is missing 'replacement_paragraphs' or 'content'; skipping")
            continue

        try:
            items = parse_patch_paragraphs(replacement_paragraphs, f"{context}.replacement_paragraphs")
        except ValueError as exc:
            warn(str(exc))
            continue

        try:
            start_idx = find_para_contains(doc, start_heading)
        except ValueError:
            warn(f"{context} start heading not found: {start_heading!r}")
            continue

        end_heading = block.get("end_heading")
        end_idx = None

        if isinstance(end_heading, str) and end_heading.strip():
            try:
                end_idx = find_para_contains(doc, end_heading, start=start_idx + 1)
            except ValueError:
                warn(f"{context} end heading not found: {end_heading!r}")
                continue
        elif require_end_heading:
            warn(f"{context} requires 'end_heading'; skipping")
            continue
        else:
            end_idx = infer_end_index(doc, start_idx)
            if end_idx is None:
                warn(f"{context} could not infer an end boundary; provide 'end_heading' to disambiguate")
                continue

        if apply_block_items(doc, start_idx, end_idx, items, context, style_templates):
            applied += 1

    return applied


def load_replacements(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"Replacements JSON not found: {path}")
    if not path.is_file():
        raise FileNotFoundError(f"Replacements path is not a file: {path}")

    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in replacements file: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError("Replacements JSON must be a top-level object")

    return data


def main() -> None:
    parser = argparse.ArgumentParser(description="Patch a resume DOCX using external replacements JSON while preserving formatting.")
    parser.add_argument("--src", required=True, help="Source DOCX resume")
    parser.add_argument("--replacements", required=True, help="Replacements JSON path")
    parser.add_argument("--out", required=True, help="Output DOCX path")
    args = parser.parse_args()

    src = Path(args.src)
    replacements_path = Path(args.replacements)
    out = Path(args.out)

    if not src.exists() or not src.is_file():
        raise FileNotFoundError(f"Source DOCX not found: {src}")

    replacements = load_replacements(replacements_path)
    doc = Document(str(src))
    if not doc.paragraphs:
        raise ValueError("Source DOCX contains no paragraphs")
    style_templates = discover_style_templates(doc)

    applied_counts: dict[str, int] = {}

    applied_counts["paragraph_replacements"] = apply_paragraph_replacements(doc, replacements.get("paragraph_replacements"))
    applied_counts["summary"] = apply_summary_replacement(doc, replacements.get("summary"))
    applied_counts["skills"] = apply_skills_replacements(doc, replacements.get("skills"))
    applied_counts["experience_blocks"] = apply_named_blocks(
        doc,
        replacements.get("experience_blocks"),
        label="experience_blocks",
        style_templates=style_templates,
        default_use_heading_key=True,
        require_end_heading=False,
    )
    applied_counts["section_replacements"] = apply_named_blocks(
        doc,
        replacements.get("section_replacements"),
        label="section_replacements",
        style_templates=style_templates,
        default_use_heading_key=True,
        require_end_heading=False,
    )
    applied_counts["block_replacements"] = apply_named_blocks(
        doc,
        replacements.get("block_replacements"),
        label="block_replacements",
        style_templates=style_templates,
        default_use_heading_key=False,
        require_end_heading=True,
    )

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))

    applied_summary = ", ".join(f"{k}={v}" for k, v in applied_counts.items())
    print(f"Wrote {out}")
    print(f"Applied operations: {applied_summary}")


if __name__ == "__main__":
    main()
