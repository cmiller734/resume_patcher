#!/usr/bin/env python3
"""
Safely patch a Word resume by replacing paragraph text while preserving the
original DOCX XML, namespaces, styles, bullets, spacing, and layout.

This script copies every file in the DOCX package unchanged except
word/document.xml, and patches only <w:t> text nodes inside selected <w:p>
paragraphs. It avoids XML reserialization because Word can flag documents as
recoverable when namespace prefixes are rewritten incorrectly.

Common usage:
  # Backward-compatible exact text replacement
  python3 -S resume_patch_docx_workhorse.py input.docx output.docx replacements.json

  # Dry run before writing anything
  python3 -S resume_patch_docx_workhorse.py input.docx output.docx replacements.json --dry-run

  # Dump paragraphs to terminal or JSON for building replacement maps
  python3 -S resume_patch_docx_workhorse.py input.docx --dump-paragraphs
  python3 -S resume_patch_docx_workhorse.py input.docx --dump-json paragraphs.json

  # Auto-generate output filename from role/company text
  python3 -S resume_patch_docx_workhorse.py input.docx replacements.json --role "Sentra Technical Support Engineer"

Replacement JSON formats supported:

1) Backward-compatible text map:
{
  "Old paragraph text": "New paragraph text",
  "Another old paragraph": "Another new paragraph"
}

2) Mixed text + index map. Numeric string keys are treated as paragraph indexes:
{
  "12": "New text for paragraph index 12",
  "Old paragraph text": "New paragraph text"
}

3) Explicit list format:
[
  {"index": 12, "replacement": "New text for paragraph index 12"},
  {"text": "Old paragraph text", "replacement": "New paragraph text"},
  {"index": 31, "expected_text": "Optional safety check", "replacement": "New text"}
]
"""

import argparse
import difflib
import html
import json
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

P_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.DOTALL)
T_RE = re.compile(r"(<w:t\b[^>]*>)(.*?)(</w:t>)", re.DOTALL)
TOKEN_RE = re.compile(r"(<w:t\b[^>]*>)(.*?)(</w:t>)|<w:br\b[^>]*/>|<w:cr\b[^>]*/>", re.DOTALL)
NUMERIC_KEY_RE = re.compile(r"^\d+$")


@dataclass
class Paragraph:
    index: int
    start: int
    end: int
    xml: str
    text: str


@dataclass
class Replacement:
    replacement: str
    source: str
    index: int | None = None
    text: str | None = None
    expected_text: str | None = None


@dataclass
class PlanResult:
    patch_by_index: dict[int, str]
    matched: list[str]
    missing: list[str]
    warnings: list[str]


def ensure_caleb_prefix(path: str | Path) -> Path:
    path = Path(path)
    if not path.name.startswith("caleb_miller_"):
        return path.with_name("caleb_miller_" + path.name)
    return path


def slugify(value: str) -> str:
    value = value.lower().strip()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "tailored_resume"


def output_from_role(input_path: str | Path, role: str) -> Path:
    input_path = Path(input_path)
    return ensure_caleb_prefix(input_path.with_name(slugify(role) + ".docx"))


def xml_unescape(s: str) -> str:
    return html.unescape(s)


def xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))


def paragraph_text(p_xml: str) -> str:
    """Return visible paragraph text, preserving explicit Word line breaks as 
.

    Word stores manual line breaks as <w:br/> or <w:cr/> rather than as
    characters inside <w:t> nodes. Preserving them here makes paragraph dumps
    and replacement matching accurately represent multi-line blocks such as
    dynamically generated Skills sections.
    """
    parts: list[str] = []
    for m in TOKEN_RE.finditer(p_xml):
        if m.group(1) is not None:
            parts.append(xml_unescape(m.group(2)))
        else:
            parts.append("\n")
    return "".join(parts)


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def text_xml_with_word_breaks(start_tag: str, end_tag: str, text: str) -> str:
    """Convert plain text containing 
 into OOXML text nodes with Word line breaks."""
    if "xml:space=" not in start_tag:
        start_tag = start_tag[:-1] + ' xml:space="preserve">'
    # Normalize Windows/Mac newlines before splitting.
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: list[str] = []
    for i, line in enumerate(lines):
        if i > 0:
            out.append("<w:br/>")
        out.append(start_tag + xml_escape(line) + end_tag)
    return "".join(out)


def set_paragraph_text(p_xml: str, new_text: str) -> str:
    matches = list(T_RE.finditer(p_xml))
    if not matches:
        return p_xml

    normalized_new_text = new_text.replace("\r\n", "\n").replace("\r", "\n")

    # Multi-line replacements need real Word line breaks. The older behavior
    # put the full replacement into the first <w:t> node and emptied the rest,
    # which flattened Skills sections and other line-separated blocks.
    # For multi-line text, replace the entire text-node span once so any old
    # <w:br/> elements between original text nodes do not create duplicate
    # or misplaced blank lines.
    if "\n" in normalized_new_text:
        first = matches[0]
        last = matches[-1]
        start_tag, _, end_tag = first.group(1), first.group(2), first.group(3)
        return (
            p_xml[:first.start()]
            + text_xml_with_word_breaks(start_tag, end_tag, normalized_new_text)
            + p_xml[last.end():]
        )

    out: list[str] = []
    last_pos = 0
    first_text = True
    for m in matches:
        out.append(p_xml[last_pos:m.start()])
        start_tag, _, end_tag = m.group(1), m.group(2), m.group(3)
        if first_text:
            if "xml:space=" not in start_tag:
                start_tag = start_tag[:-1] + ' xml:space="preserve">'
            out.append(start_tag + xml_escape(normalized_new_text) + end_tag)
            first_text = False
        else:
            out.append(start_tag + "" + end_tag)
        last_pos = m.end()
    out.append(p_xml[last_pos:])
    return "".join(out)


def read_document_xml(input_path: str | Path) -> str:
    with zipfile.ZipFile(input_path, "r") as zin:
        return zin.read("word/document.xml").decode("utf-8")


def extract_paragraphs(doc_xml: str) -> list[Paragraph]:
    paragraphs: list[Paragraph] = []
    for i, match in enumerate(P_RE.finditer(doc_xml)):
        p_xml = match.group(0)
        paragraphs.append(
            Paragraph(
                index=i,
                start=match.start(),
                end=match.end(),
                xml=p_xml,
                text=paragraph_text(p_xml).strip(),
            )
        )
    return paragraphs


def load_replacements(path: str | Path) -> list[Replacement]:
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    replacements: list[Replacement] = []

    if isinstance(raw, dict):
        for key, value in raw.items():
            if not isinstance(value, str):
                raise SystemExit(f"Replacement value for key {key!r} must be a string.")
            if NUMERIC_KEY_RE.match(str(key)):
                replacements.append(Replacement(index=int(key), replacement=value, source=f"index:{key}"))
            else:
                replacements.append(Replacement(text=str(key), replacement=value, source=f"text:{str(key)[:60]}"))
        return replacements

    if isinstance(raw, list):
        for idx, item in enumerate(raw):
            if not isinstance(item, dict):
                raise SystemExit(f"Replacement item #{idx} must be an object.")
            if "replacement" not in item or not isinstance(item["replacement"], str):
                raise SystemExit(f"Replacement item #{idx} must include a string 'replacement'.")
            has_index = "index" in item
            has_text = "text" in item
            if has_index == has_text:
                raise SystemExit(f"Replacement item #{idx} must include exactly one of 'index' or 'text'.")
            expected_text = item.get("expected_text")
            if expected_text is not None and not isinstance(expected_text, str):
                raise SystemExit(f"Replacement item #{idx} has non-string expected_text.")
            if has_index:
                replacements.append(
                    Replacement(
                        index=int(item["index"]),
                        expected_text=expected_text,
                        replacement=item["replacement"],
                        source=f"item#{idx}:index:{item['index']}",
                    )
                )
            else:
                replacements.append(
                    Replacement(
                        text=str(item["text"]),
                        replacement=item["replacement"],
                        source=f"item#{idx}:text:{str(item['text'])[:60]}",
                    )
                )
        return replacements

    raise SystemExit("Replacement JSON must be either an object or a list of replacement objects.")


def find_text_matches(
    paragraphs: list[Paragraph],
    needle: str,
    *,
    contains_match: bool,
    fuzzy_threshold: float | None,
) -> tuple[list[int], str]:
    # 1. Exact match preserves the original script's safest behavior.
    exact = [p.index for p in paragraphs if p.text == needle]
    if exact:
        return exact, "exact"

    # 2. Normalized whitespace match handles Word line/space quirks without being too loose.
    normalized_needle = normalize_text(needle)
    normalized = [p.index for p in paragraphs if normalize_text(p.text) == normalized_needle]
    if normalized:
        return normalized, "normalized"

    # 3. Optional contains match. Only useful when source snippets are intentionally partial.
    if contains_match:
        contains = [p.index for p in paragraphs if normalized_needle in normalize_text(p.text)]
        if contains:
            return contains, "contains"

    # 4. Optional fuzzy match. Require caller opt-in.
    if fuzzy_threshold is not None:
        scored = []
        for p in paragraphs:
            ratio = difflib.SequenceMatcher(None, normalized_needle, normalize_text(p.text)).ratio()
            if ratio >= fuzzy_threshold:
                scored.append((ratio, p.index))
        scored.sort(reverse=True)
        if scored:
            best_ratio = scored[0][0]
            best = [index for ratio, index in scored if ratio == best_ratio]
            return best, f"fuzzy:{best_ratio:.3f}"

    return [], "missing"


def build_patch_plan(
    paragraphs: list[Paragraph],
    replacements: list[Replacement],
    *,
    contains_match: bool = False,
    fuzzy_threshold: float | None = None,
) -> PlanResult:
    patch_by_index: dict[int, str] = {}
    matched: list[str] = []
    missing: list[str] = []
    warnings: list[str] = []

    by_index = {p.index: p for p in paragraphs}

    for repl in replacements:
        if repl.index is not None:
            if repl.index not in by_index:
                missing.append(f"{repl.source} - paragraph index does not exist")
                continue
            paragraph = by_index[repl.index]
            if repl.expected_text is not None:
                expected = normalize_text(repl.expected_text)
                actual = normalize_text(paragraph.text)
                if expected != actual:
                    warnings.append(
                        f"{repl.source} - expected_text did not match paragraph text. "
                        "Patch still planned because index was explicit."
                    )
            if repl.index in patch_by_index:
                warnings.append(f"{repl.source} - duplicate replacement for paragraph {repl.index}; later value wins.")
            patch_by_index[repl.index] = repl.replacement
            matched.append(f"{repl.source} -> paragraph {repl.index} (index)")
            continue

        assert repl.text is not None
        matches, mode = find_text_matches(
            paragraphs,
            repl.text,
            contains_match=contains_match,
            fuzzy_threshold=fuzzy_threshold,
        )
        if not matches:
            missing.append(f"{repl.source} - no matching paragraph found")
            continue
        if len(matches) > 1 and mode in {"contains"} or (len(matches) > 1 and mode.startswith("fuzzy")):
            missing.append(
                f"{repl.source} - ambiguous {mode} match across paragraphs {matches}. "
                "Use an index-based replacement for this one."
            )
            continue
        for paragraph_index in matches:
            if paragraph_index in patch_by_index:
                warnings.append(
                    f"{repl.source} - paragraph {paragraph_index} already has a planned replacement; later value wins."
                )
            patch_by_index[paragraph_index] = repl.replacement
        matched.append(f"{repl.source} -> paragraph(s) {matches} ({mode})")

    return PlanResult(patch_by_index=patch_by_index, matched=matched, missing=missing, warnings=warnings)


def apply_patch_plan(doc_xml: str, paragraphs: list[Paragraph], patch_by_index: dict[int, str]) -> str:
    if not patch_by_index:
        return doc_xml

    out: list[str] = []
    last = 0
    for paragraph in paragraphs:
        if paragraph.index not in patch_by_index:
            continue
        out.append(doc_xml[last:paragraph.start])
        out.append(set_paragraph_text(paragraph.xml, patch_by_index[paragraph.index]))
        last = paragraph.end
    out.append(doc_xml[last:])
    return "".join(out)


def write_patched_docx(input_path: str | Path, output_path: str | Path, patched_xml: str) -> Path:
    input_path = Path(input_path)
    output_path = ensure_caleb_prefix(output_path)
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = patched_xml.encode("utf-8")
                zout.writestr(item, data)
    return output_path


def validate_docx(path: str | Path) -> list[str]:
    warnings: list[str] = []
    try:
        doc_xml = read_document_xml(path)
    except Exception as exc:  # noqa: BLE001 - this is a user-facing sanity check
        return [f"Could not read word/document.xml from output DOCX: {exc}"]

    text = "\n".join(p.text for p in extract_paragraphs(doc_xml) if p.text)
    required_terms = ["Caleb Miller", "SUMMARY", "SKILLS", "EDUCATION"]
    for term in required_terms:
        if term not in text:
            warnings.append(f"Missing expected resume term/section: {term}")

    if "WORK EXPERIENCE" not in text and "Work EXPERIENCE" not in text:
        warnings.append("Missing expected resume section: WORK EXPERIENCE")

    placeholder_patterns = [r"\bTODO\b", r"\bTK\b", r"PLACEHOLDER", r"\[.*?\]"]
    for pattern in placeholder_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            warnings.append(f"Possible placeholder text found matching pattern: {pattern}")

    return warnings


def print_plan_report(plan: PlanResult, *, dry_run: bool) -> None:
    print("Patch plan:")
    print(f"- Planned paragraph replacements: {len(plan.patch_by_index)}")
    print(f"- Matched replacement requests: {len(plan.matched)}")
    print(f"- Missing/blocked replacement requests: {len(plan.missing)}")
    print(f"- Warnings: {len(plan.warnings)}")

    if plan.matched:
        print("\nMatched:")
        for item in plan.matched:
            print(f"- {item}")

    if plan.warnings:
        print("\nWarnings:")
        for item in plan.warnings:
            print(f"- {item}")

    if plan.missing:
        print("\nMissing/blocked:")
        for item in plan.missing:
            print(f"- {item}")

    if dry_run:
        print("\nDry run only. No output file was written.")


def dump_paragraphs(input_path: str | Path) -> None:
    paragraphs = extract_paragraphs(read_document_xml(input_path))
    for p in paragraphs:
        if p.text:
            print(f"[{p.index}]\n{p.text}\n")


def dump_json(input_path: str | Path, output_path: str | Path) -> Path:
    paragraphs = extract_paragraphs(read_document_xml(input_path))
    payload = [
        {"index": p.index, "text": p.text}
        for p in paragraphs
        if p.text
    ]
    output_path = Path(output_path)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return output_path


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Safely patch DOCX resume text without XML namespace reserialization.")
    parser.add_argument("input", help="Input .docx file")
    parser.add_argument(
        "positional",
        nargs="*",
        help=(
            "Either output.docx replacements.json, or replacements.json when --role is used. "
            "Legacy usage is preserved."
        ),
    )
    parser.add_argument("--dump-paragraphs", action="store_true", help="Print non-empty paragraph indexes and text.")
    parser.add_argument("--dump-json", help="Write non-empty paragraph indexes and text to a JSON file.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be patched without writing an output file.")
    parser.add_argument("--allow-missing", action="store_true", help="Patch matched paragraphs even if some requests are missing.")
    parser.add_argument("--contains-match", action="store_true", help="Allow text keys to match paragraphs containing that text.")
    parser.add_argument(
        "--fuzzy-threshold",
        type=float,
        default=None,
        help="Optional fuzzy match threshold from 0.0 to 1.0. Suggested: 0.92+. Use carefully.",
    )
    parser.add_argument("--role", help="Auto-generate output filename from role/company text.")
    parser.add_argument("--strict-validation", action="store_true", help="Treat validation warnings as fatal.")
    return parser.parse_args(argv)


def resolve_output_and_replacements(args: argparse.Namespace) -> tuple[Path, Path]:
    positional = [Path(p) for p in args.positional]

    if args.role:
        if len(positional) != 1:
            raise SystemExit("When using --role, provide exactly one positional argument: replacements.json")
        return output_from_role(args.input, args.role), positional[0]

    if len(positional) != 2:
        raise SystemExit("Expected output.docx and replacements.json unless using --dump-paragraphs, --dump-json, or --role.")

    return ensure_caleb_prefix(positional[0]), positional[1]


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)

    if args.fuzzy_threshold is not None and not (0 <= args.fuzzy_threshold <= 1):
        raise SystemExit("--fuzzy-threshold must be between 0.0 and 1.0")

    if args.dump_paragraphs:
        dump_paragraphs(args.input)
        return 0

    if args.dump_json:
        path = dump_json(args.input, args.dump_json)
        print(f"Wrote paragraph JSON: {path}")
        return 0

    output_path, replacements_path = resolve_output_and_replacements(args)
    doc_xml = read_document_xml(args.input)
    paragraphs = extract_paragraphs(doc_xml)
    replacements = load_replacements(replacements_path)
    plan = build_patch_plan(
        paragraphs,
        replacements,
        contains_match=args.contains_match,
        fuzzy_threshold=args.fuzzy_threshold,
    )

    print_plan_report(plan, dry_run=args.dry_run)

    if plan.missing and not args.allow_missing:
        raise SystemExit("\nAborted because at least one replacement was missing or ambiguous. Use --allow-missing to patch matches anyway.")

    if args.dry_run:
        return 0

    patched_xml = apply_patch_plan(doc_xml, paragraphs, plan.patch_by_index)
    written_path = write_patched_docx(args.input, output_path, patched_xml)
    print(f"\nWrote patched DOCX: {written_path}")

    validation_warnings = validate_docx(written_path)
    if validation_warnings:
        print("\nValidation warnings:")
        for warning in validation_warnings:
            print(f"- {warning}")
        if args.strict_validation:
            raise SystemExit("Strict validation failed.")
    else:
        print("Validation passed.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
