#!/usr/bin/env python3
"""Regression check for Work Experience semantic paragraph styles."""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()


def paragraph_style_id(paragraph: ET.Element) -> str:
    pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
    if pstyle is None:
        return ""
    return pstyle.attrib.get(f"{{{W}}}val", "")


def paragraph_has_numbering(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:numPr", NS) is not None


def paragraph_space_after(paragraph: ET.Element) -> int:
    spacing = paragraph.find("./w:pPr/w:spacing", NS)
    if spacing is None:
        return 0
    raw = spacing.attrib.get(f"{{{W}}}after")
    return int(raw) if raw is not None else 0


def paragraph_runs(paragraph: ET.Element) -> list[ET.Element]:
    return paragraph.findall(".//w:r", NS)


def run_text(run: ET.Element) -> str:
    return "".join(t.text or "" for t in run.findall(".//w:t", NS))


def run_bool(run: ET.Element, tag_name: str) -> bool | None:
    element = run.find(f"./w:rPr/w:{tag_name}", NS)
    if element is None:
        return None
    value = element.attrib.get(f"{{{W}}}val")
    return value not in {"0", "false", "False", "off"}


def run_format(run: ET.Element) -> tuple[bool | None, bool | None]:
    return run_bool(run, "b"), run_bool(run, "i")


def paragraph_line_runs(paragraph: ET.Element) -> list[list[ET.Element]]:
    lines: list[list[ET.Element]] = [[]]
    for run in paragraph_runs(paragraph):
        has_visible_text_on_line = False
        for child in run:
            if child.tag == f"{{{W}}}t" and (child.text or "").strip():
                has_visible_text_on_line = True
            elif child.tag in {f"{{{W}}}br", f"{{{W}}}cr"}:
                if has_visible_text_on_line:
                    lines[-1].append(run)
                lines.append([])
                has_visible_text_on_line = False
        if has_visible_text_on_line:
            lines[-1].append(run)
    return lines


def first_visible_run_format(paragraph: ET.Element, line_idx: int = 0) -> tuple[bool | None, bool | None]:
    lines = paragraph_line_runs(paragraph)
    candidates = lines[min(line_idx, len(lines) - 1)] if lines else []
    for run in candidates:
        if run_text(run).strip():
            return run_format(run)
    raise AssertionError(f"No visible run found on line {line_idx} of {paragraph_text(paragraph)!r}")


def find_paragraph_containing(paragraphs: list[ET.Element], *needles: str) -> ET.Element:
    for paragraph in paragraphs:
        text = paragraph_text(paragraph)
        if all(needle in text for needle in needles):
            return paragraph
    raise AssertionError(f"Could not find paragraph containing: {needles!r}")


def read_docx_paragraphs(path: Path) -> tuple[list[ET.Element], dict[str, str]]:
    with ZipFile(path) as zf:
        document_xml = ET.fromstring(zf.read("word/document.xml"))
        styles_xml = ET.fromstring(zf.read("word/styles.xml"))

    style_names: dict[str, str] = {}
    for style in styles_xml.findall(".//w:style", NS):
        style_id = style.attrib.get(f"{{{W}}}styleId")
        name = style.find("w:name", NS)
        if style_id and name is not None:
            style_names[style_id] = name.attrib.get(f"{{{W}}}val", style_id)

    return document_xml.findall(".//w:body/w:p", NS), style_names


def find_paragraph(paragraphs: list[ET.Element], exact_text: str) -> ET.Element:
    for paragraph in paragraphs:
        if paragraph_text(paragraph) == exact_text:
            return paragraph
    raise AssertionError(f"Could not find generated paragraph: {exact_text!r}")


def assert_not_heading_1(paragraph: ET.Element, style_names: dict[str, str], label: str) -> None:
    style_id = paragraph_style_id(paragraph)
    style_name = style_names.get(style_id, style_id).lower().replace(" ", "")
    if style_id.lower() == "heading1" or style_name == "heading1":
        raise AssertionError(f"{label} unexpectedly received Heading 1")


def assert_heading_1(paragraph: ET.Element, style_names: dict[str, str], label: str) -> None:
    assert_style(paragraph, style_names, "heading1", label)


def assert_style(paragraph: ET.Element, style_names: dict[str, str], expected_style: str, label: str) -> None:
    style_id = paragraph_style_id(paragraph)
    style_name = style_names.get(style_id, style_id).lower().replace(" ", "")
    expected = expected_style.lower().replace(" ", "")
    if style_id.lower() != expected and style_name != expected:
        raise AssertionError(f"{label} did not keep {expected_style}; got {style_id!r}/{style_name!r}")


def assert_runs_match_format(paragraph: ET.Element, label: str, expected: tuple[bool | None, bool | None]) -> None:
    for run in paragraph_runs(paragraph):
        if not run_text(run).strip():
            continue
        actual = run_format(run)
        if actual != expected:
            raise AssertionError(f"{label} run {run_text(run)!r} format={actual}, expected {expected}")


def assert_bullet(paragraph: ET.Element, style_names: dict[str, str], label: str) -> None:
    style_id = paragraph_style_id(paragraph)
    style_name = style_names.get(style_id, style_id).lower()
    if not paragraph_has_numbering(paragraph) and "bullet" not in style_name and "list paragraph" not in style_name:
        raise AssertionError(f"{label} did not keep bullet/list formatting; got {style_id!r}/{style_name!r}")


def find_bullet_template(paragraphs: list[ET.Element], style_names: dict[str, str]) -> ET.Element:
    for paragraph in paragraphs:
        text = paragraph_text(paragraph)
        if not text:
            continue
        style_id = paragraph_style_id(paragraph)
        style_name = style_names.get(style_id, style_id).lower()
        if paragraph_has_numbering(paragraph) or "bullet" in style_name or "list paragraph" in style_name:
            return paragraph
    raise AssertionError("Could not find a source bullet paragraph to use as the master template")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--src", default="Caleb Miller Master Resume.docx", help="Source DOCX")
    parser.add_argument("--out", default="", help="Optional output DOCX path")
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parents[1]
    src = Path(args.src)
    if not src.is_absolute():
        src = repo_root / src

    replacement = {
        "skills": {
            "Support & Platform": "regression operating systems, escalation handling, documentation"
        },
        "block_replacements": [
            {
                "start_heading": "Work EXPERIENCE",
                "end_heading": "EDUCATION",
                "content": [
                    {"text": "Work EXPERIENCE", "style": "section_heading"},
                    {
                        "text": "Regression Company | Enterprise SaaS Application Support",
                        "style": "company_heading",
                    },
                    {"text": "Professional Services Engineer", "style": "role_heading"},
                    {"text": "Remote | Mar. 2024 - Feb. 2025", "style": "date_location"},
                    {
                        "text": "Verified that inserted bullets keep list formatting instead of becoming headings.",
                        "style": "bullet",
                    },
                    {
                        "text": "Second Regression Company | Application Operations",
                        "style": "company_heading",
                    },
                    {"text": "Application Support Engineer", "style": "role_heading"},
                    {"text": "Remote | Jan. 2023 - Feb. 2024", "style": "date_location"},
                    {
                        "text": "Verified that generated experience items keep readable spacing after each item.",
                        "style": "bullet",
                    },
                ],
            }
        ]
    }

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        replacements_path = tmp_path / "replacements.json"
        replacements_path.write_text(json.dumps(replacement, indent=2), encoding="utf-8")
        out = Path(args.out) if args.out else tmp_path / "tailored_resume.docx"

        subprocess.run(
            [
                sys.executable,
                str(repo_root / "resume_patcher.py"),
                "--src",
                str(src),
                "--replacements",
                str(replacements_path),
                "--out",
                str(out),
            ],
            check=True,
        )

        source_paragraphs, source_style_names = read_docx_paragraphs(src)
        paragraphs, style_names = read_docx_paragraphs(out)

        for section_heading in {"SUMMARY", "SKILLS", "Work EXPERIENCE", "EDUCATION"}:
            source_heading = find_paragraph(source_paragraphs, section_heading)
            output_heading = find_paragraph(paragraphs, section_heading)
            assert_style(
                output_heading,
                style_names,
                paragraph_style_id(source_heading) or source_style_names.get(paragraph_style_id(source_heading), ""),
                f"{section_heading} section",
            )

        source_role_date = find_paragraph_containing(
            source_paragraphs,
            "Professional Services Engineer",
            "Remote | Mar. 2024 - Feb. 2025",
        )
        expected_header_style = paragraph_style_id(source_role_date)
        expected_role_format = first_visible_run_format(source_role_date, 0)
        expected_date_format = first_visible_run_format(source_role_date, 1)

        role_heading = find_paragraph(paragraphs, "Professional Services Engineer")
        date_location = find_paragraph(paragraphs, "Remote | Mar. 2024 - Feb. 2025")
        assert_style(role_heading, style_names, expected_header_style, "role heading")
        assert_runs_match_format(role_heading, "role heading", expected_role_format)
        assert_style(date_location, style_names, expected_header_style, "date/location")
        assert_runs_match_format(date_location, "date/location", expected_date_format)

        assert_not_heading_1(role_heading, style_names, "role heading")
        assert_not_heading_1(date_location, style_names, "date/location")

        company_heading = find_paragraph(paragraphs, "Regression Company | Enterprise SaaS Application Support")
        assert_not_heading_1(company_heading, style_names, "company heading")

        bullet = find_paragraph(
            paragraphs,
            "Verified that inserted bullets keep list formatting instead of becoming headings.",
        )
        assert_bullet(
            bullet,
            style_names,
            "bullet",
        )
        source_bullet = find_bullet_template(source_paragraphs, source_style_names)
        assert_runs_match_format(bullet, "bullet", first_visible_run_format(source_bullet))

        second_bullet = find_paragraph(
            paragraphs,
            "Verified that generated experience items keep readable spacing after each item.",
        )
        expected_spacing_after = 120
        if paragraph_space_after(bullet) < expected_spacing_after:
            raise AssertionError(
                f"first generated experience item did not keep trailing spacing: {paragraph_space_after(bullet)}"
            )
        if paragraph_space_after(second_bullet) < expected_spacing_after:
            raise AssertionError(
                f"final generated experience item did not keep trailing spacing: {paragraph_space_after(second_bullet)}"
            )

        skills = find_paragraph(
            paragraphs,
            "Support & Platform: regression operating systems, escalation handling, documentation",
        )
        source_skills = find_paragraph_containing(source_paragraphs, "Support & Platform:")
        source_skill_runs = paragraph_runs(source_skills)
        source_label_format = run_format(next(run for run in source_skill_runs if run_text(run).startswith("Support & Platform:")))
        source_rest_format = run_format(next(run for run in source_skill_runs if "technical support" in run_text(run)))

        skills_runs = [(run_text(run), run_format(run)) for run in paragraph_runs(skills) if run_text(run)]
        if ("Support & Platform:", source_label_format) not in skills_runs:
            raise AssertionError(f"skills label did not copy master formatting: {skills_runs!r}")
        if not any("regression operating systems" in text and fmt == source_rest_format for text, fmt in skills_runs):
            raise AssertionError(f"skills text after colon did not copy master formatting: {skills_runs!r}")

    print("OK: Work Experience and Skills formatting copied the master templates.")


if __name__ == "__main__":
    main()
