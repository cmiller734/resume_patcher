#!/usr/bin/env python3
"""
resume_patch_docx_workhorse_linebreaks.py

Patch an existing DOCX resume in-place-style rather than rebuilding from scratch.
Designed for Caleb's resume workflow:
- preserve the source document's section/layout feel
- replace targeted paragraph blocks by anchor text
- preserve run formatting and paragraph styles
- support multi-line header paragraphs without losing line breaks
"""
from __future__ import annotations

import argparse
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence

from docx import Document
from docx.text.paragraph import Paragraph


@dataclass(frozen=True)
class PatchPara:
    text: str
    style_source: str = "normal"  # normal | bullet | keep


def _copy_run_format(src_run, dst_run) -> None:
    """Copy formatting from src_run to dst_run without copying text."""
    if src_run is None:
        return
    if src_run._r.rPr is not None:
        dst_run._r.insert(0, deepcopy(src_run._r.rPr))


def set_paragraph_text_keep_format(paragraph: Paragraph, text: str, template: Paragraph | None = None) -> None:
    """Replace paragraph text while preserving paragraph style and first-run formatting.

    Uses add_break() for embedded newlines so header/date/progression lines stay in one
    paragraph exactly like the original resume's combined header paragraphs.
    """
    template = template or paragraph
    src_run = template.runs[0] if template is not None and template.runs else (paragraph.runs[0] if paragraph.runs else None)

    # Remove existing runs.
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)

    parts = text.split("\n")
    run = paragraph.add_run(parts[0] if parts else "")
    _copy_run_format(src_run, run)
    for part in parts[1:]:
        run.add_break()
        run.add_text(part)


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


def find_para_contains(doc: Document, needle: str, start: int = 0) -> int:
    for i, p in enumerate(doc.paragraphs[start:], start=start):
        if needle in p.text:
            return i
    raise ValueError(f"Could not find paragraph containing: {needle!r}")


def replace_block(doc: Document, start_idx: int, end_idx_exclusive: int, items: Sequence[PatchPara], *, normal_template: Paragraph, bullet_template: Paragraph) -> None:
    """Replace a paragraph block [start_idx, end_idx_exclusive) with items.

    Existing content is removed and new paragraphs are inserted using templates from the
    source document. This preserves the visual system while making targeted edits.
    The end paragraph is treated as an anchor and is not deleted.
    """
    if not items:
        raise ValueError("replace_block requires at least one item")

    first = doc.paragraphs[start_idx]
    end_anchor = doc.paragraphs[end_idx_exclusive]._p

    # Preserve template XML before any deletion, so deleted template paragraphs can still
    # be safely cloned for inserted content.
    normal_template_xml = deepcopy(normal_template._p)
    bullet_template_xml = deepcopy(bullet_template._p)

    def make_template(style_source: str, current: Paragraph | None = None) -> Paragraph:
        if style_source == "bullet":
            return Paragraph(deepcopy(bullet_template_xml), first._parent)
        if style_source == "normal":
            return Paragraph(deepcopy(normal_template_xml), first._parent)
        return current or Paragraph(deepcopy(normal_template_xml), first._parent)

    first.style = make_template(items[0].style_source, first).style
    set_paragraph_text_keep_format(first, items[0].text, make_template(items[0].style_source, first))

    # Delete everything after first until the preserved end anchor. This is safer than
    # relying on counts when Word inserts or preserves extra blank paragraphs.
    while first._p.getnext() is not None and first._p.getnext() is not end_anchor:
        delete_paragraph(Paragraph(first._p.getnext(), first._parent))

    cursor = first
    for item in items[1:]:
        cursor = insert_paragraph_after(cursor, item.text, make_template(item.style_source))


def main() -> None:
    parser = argparse.ArgumentParser(description="Patch Caleb Miller resume DOCX while preserving formatting.")
    parser.add_argument("--src", required=True, help="Source DOCX resume")
    parser.add_argument("--out", required=True, help="Output DOCX path")
    args = parser.parse_args()

    src = Path(args.src)
    out = Path(args.out)
    doc = Document(str(src))

    # Source style templates from the original document.
    normal_template = doc.paragraphs[9]   # role/header paragraph style/run formatting
    bullet_template = doc.paragraphs[10]  # list paragraph bullet style/run formatting

    # SUMMARY
    summary_idx = find_para_index(doc, "Application systems and infrastructure specialist with 9 years of experience supporting enterprise SaaS platforms, production web applications, and business systems environments. Background in technical support, product implementation, and application operations across Windows and Linux server environments. Experienced with API integrations, SQL scripting, SSL/TLS certificate management, and maintaining reliable production systems. Known for bridging development and operations to deploy, maintain, and troubleshoot complex application environments.")
    set_paragraph_text_keep_format(
        doc.paragraphs[summary_idx],
        "Customer Success Engineer and technical implementation specialist with 9 years of experience across enterprise SaaS support, product implementation, application operations, and customer-facing workflow configuration. Strong at translating complex customer requirements into structured systems, configuring production environments, documenting repeatable rollout processes, and using SQL, APIs, logs, and telemetry to investigate issues. Background spans frontline support, Professional Services, customer training, and product feedback loops in small, high-ownership SaaS teams.",
        doc.paragraphs[summary_idx],
    )

    # EXPERIENCE block: replace only the main/current experience block.
    # Do NOT generate or collapse the older-development section here.
    # The master resume already contains an "Additional Web Development Experience"
    # section, and this script should preserve that existing source section/formatting.
    exp_start = find_para_contains(doc, "Freelance | WordPress Developer", start=0)
    try:
        exp_end_idx = find_para_contains(doc, "Additional Web Development Experience", start=exp_start)
    except ValueError:
        exp_end_idx = find_para_index(doc, "EDUCATION")

    experience_items: List[PatchPara] = [
        PatchPara("Apica | Enterprise SaaS Support, Implementation & Professional Services\nPortsmouth, NH / Remote | Mar. 2019 - Feb. 2025\nProgression: Technical Support Engineer -> Senior Technical Support Engineer -> Professional Services Engineer", "normal"),
        PatchPara("Advanced from Technical Support Engineer to Senior Technical Support Engineer to Professional Services Engineer, taking on implementation ownership, enterprise escalation support, customer training, and billable Professional Services work.", "bullet"),
        PatchPara("Acted as primary U.S. technical resource for enterprise customers, balancing responsive support, implementation work, configuration guidance, and long-term adoption needs.", "bullet"),
        PatchPara("Led customer implementations of Apica's observability platform, configuring SaaS and on-premise environments to match customer infrastructure, monitoring goals, reporting needs, and operational constraints.", "bullet"),
        PatchPara("Used SQL queries, REST APIs, Postman, logs, and platform telemetry to investigate production issues, validate configurations, and support customer-specific data workflows.", "bullet"),
        PatchPara("Built dashboards, reporting views, and custom technical solutions to help customers interpret system behavior, monitor uptime, and reduce manual effort.", "bullet"),
        PatchPara("Captured recurring customer patterns, edge cases, and product behavior questions, translating them into documentation updates, bug reports, and actionable Product/Engineering follow-up.", "bullet"),
        PatchPara("Created and maintained customer-facing and internal documentation covering deployments, product behavior, support procedures, and repeatable implementation steps.", "bullet"),
        PatchPara("Led customer training and support conversations over tickets, remote sessions, and meetings while maintaining strong SLA discipline and customer trust.", "bullet"),
        PatchPara("", "normal"),
        PatchPara("Freelance | WordPress Developer & Systems Architect\nRemote | June 2025 - Present", "normal"),
        PatchPara("Built and maintained production WordPress/WooCommerce systems combining backend business logic, structured configuration, customer-facing workflows, and reporting needs.", "bullet"),
        PatchPara("Architected and maintained a commercial WooCommerce extension for subscription changes and box configuration, owning business logic, debugging, testing, release validation, and customer-facing setup.", "bullet"),
        PatchPara("Worked directly with business owners to translate manual operational processes into maintainable application workflows, configuration rules, documentation, and training.", "bullet"),
        PatchPara("Managed DNS, SSL/TLS, updates, caching, and performance optimization across hosted production systems.", "bullet"),
        PatchPara("", "normal"),
        PatchPara("Pearl Marketing | WordPress Developer & Systems Architect (Contract)\nRemote | May 2025 - Present", "normal"),
        PatchPara("Reworked a complex WooCommerce booking flow to make date, product, and operational rules easier to manage and less fragile for non-technical users.", "bullet"),
        PatchPara("Partnered with business stakeholders to scope requirements, validate edge cases, and deliver configuration-driven solutions aligned with real-world operations.", "bullet"),
        PatchPara("Troubleshot production issues across client environments, including plugin behavior, forms/workflows, front-end issues, performance, and usability.", "bullet"),
        PatchPara("", "normal"),
    ]
    replace_block(doc, exp_start, exp_end_idx, experience_items, normal_template=normal_template, bullet_template=bullet_template)

    # SKILLS: replace three skill rows while preserving the original section position.
    # Re-resolve after block replacement.
    skill_start = find_para_contains(doc, "Systems & Infrastructure:")
    skills = [
        "Customer Success & Implementation: SaaS implementation, customer onboarding, customer training, workflow configuration, forms/templates, technical documentation, support ownership, JIRA",
        "Systems, Data & Troubleshooting: SQL, REST APIs, Postman, logs, structured data, reporting dashboards, Windows Server, Linux, DNS, SSL/TLS, IIS, Apache Tomcat",
        "Application & Web: WordPress, WooCommerce, PHP, JavaScript, HTML/CSS, Git, production application support, configuration-driven workflows",
    ]
    for offset, text in enumerate(skills):
        set_paragraph_text_keep_format(doc.paragraphs[skill_start + offset], text, doc.paragraphs[skill_start + offset])

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
