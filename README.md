# Resume DOCX Patcher

A conservative `.docx` patching utility for tailoring resumes without destroying the original Word formatting.

This script replaces paragraph text inside an existing Word document while preserving the original `.docx` package, styles, bullets, spacing, layout, and XML structure as much as possible.

It is designed for resume tailoring workflows where the source document is already formatted correctly and the goal is to make targeted content edits quickly.



## Why this exists

Tailoring resumes manually is slow, but rebuilding `.docx` files programmatically can easily create formatting problems:

- changed fonts or spacing
- broken bullets
- weird indentation
- flattened line breaks
- inconsistent section formatting
- Word “repair” warnings
- damaged XML namespace handling

This script avoids those problems by copying the original `.docx` package unchanged and patching only selected text nodes inside `word/document.xml`.

In plain English:

> It patches the existing resume instead of recreating it.

## What it does

The patcher can:

- replace resume paragraphs by exact text
- replace resume paragraphs by paragraph index
- preserve Word formatting, paragraph styles, bullets, and spacing
- preserve multi-line replacements using real Word line breaks
- dump paragraph indexes to the terminal or JSON
- run a dry run before writing output
- auto-generate output filenames from a role/company name
- validate the resulting document for expected resume sections
- avoid XML reserialization that can cause Word repair errors

## What it does not do

This tool is intentionally narrow.

It does not:

- redesign resumes
- create resumes from scratch
- rewrite content automatically
- edit headers, footers, styles, images, or comments
- understand resume semantics on its own
- guarantee the final resume is visually perfect without review

You should still open the final `.docx` in Microsoft Word, LibreOffice, or another Word-compatible editor before sending it.

## Requirements

- Python 3.10+
- No third-party dependencies

The script uses only the Python standard library.

## Installation

Clone the repository or download the script directly.

```bash
git clone https://github.com/your-username/resume-docx-patcher.git
cd resume-docx-patcher