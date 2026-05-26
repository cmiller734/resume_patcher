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
    style_id = paragraph_style_id(paragraph)
    style_name = style_names.get(style_id, style_id).lower().replace(" ", "")
    if style_id.lower() != "heading1" and style_name != "heading1":
        raise AssertionError(f"{label} did not keep Heading 1; got {style_id!r}/{style_name!r}")


def assert_bullet(paragraph: ET.Element, style_names: dict[str, str], label: str) -> None:
    style_id = paragraph_style_id(paragraph)
    style_name = style_names.get(style_id, style_id).lower()
    if not paragraph_has_numbering(paragraph) and "bullet" not in style_name and "list paragraph" not in style_name:
        raise AssertionError(f"{label} did not keep bullet/list formatting; got {style_id!r}/{style_name!r}")


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

        paragraphs, style_names = read_docx_paragraphs(out)
        assert_heading_1(find_paragraph(paragraphs, "Work EXPERIENCE"), style_names, "section heading")
        assert_not_heading_1(
            find_paragraph(paragraphs, "Regression Company | Enterprise SaaS Application Support"),
            style_names,
            "company heading",
        )
        assert_not_heading_1(find_paragraph(paragraphs, "Professional Services Engineer"), style_names, "role heading")
        assert_not_heading_1(
            find_paragraph(paragraphs, "Remote | Mar. 2024 - Feb. 2025"),
            style_names,
            "date/location",
        )
        assert_bullet(
            find_paragraph(
                paragraphs,
                "Verified that inserted bullets keep list formatting instead of becoming headings.",
            ),
            style_names,
            "bullet",
        )

    print("OK: Work Experience semantic styles did not inherit Heading 1.")


if __name__ == "__main__":
    main()
