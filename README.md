# AI Resume Tailoring App

An AI-assisted resume tailoring workflow that combines chatbot judgment with deterministic DOCX patching.

This repository is designed to be used with an AI chatbot. The chatbot analyzes a target job description and generates substitution/replacement content. If `resume_preferences.json` is present, the chatbot uses it by default unless the user asks not to. Then `resume_patcher.py` performs the document conversion step by applying updates to a source `.docx` while preserving formatting patterns.

## Overview

This is not designed to be a "resume generation pipeline". The goal of this open-source project is to utilize the latest breakthroughs in technology to allow real people to succeed in their careers. It also helps recruiters and hiring managers by giving them a very well-structured document with which to evaluate their candidates.

This is also not designed to cheat the system. If you make up experience on your resume or claim to have experience you don't actually have, you will be wasting both your time and your recruiter's time. It IS designed to help you present your unique, human, hard-won experience in a way that is helpful for both you AND hiring managers.

## How To Use

1. Clone the repository to your system.
2. Add your master resume to the repository. Include everything relevant you have done professionally in your field in that master resume.
3. Zip the repository and upload it to the AI chatbot of your choice.
4. Paste the job description into the chat prompt or upload it as a separate `.txt` file.
5. Ask the chatbot to use the attached ZIP and enclosed instructions to create your tailored resume.
   - Note: if `resume_preferences.json` is present, the chatbot should use it by default unless you ask it not to.
6. Work with the finished resume and the chatbot to perform final edits. This tool does not eliminate human judgment or final review; the resume still needs to impress an actual human.

## Repository Contents

- `resume_patcher.py`: Deterministic DOCX conversion/patching engine.
- `resume_preferences.json` (optional): User preference profile for AI generation decisions (tone, bullet density, framing, ATS balance, credibility guardrails).
- `requirements.txt`: Python dependency list for the patcher.

## App Architecture

This workflow intentionally separates responsibilities:

1. AI Layer (chatbot):
- Reads the job description and source resume context.
- Uses `resume_preferences.json` by default when present, unless the user asks to ignore it for that run.
- Produces proposed substitutions/revisions (optionally as structured JSON for review).
- Creates `replacements.json` before any patching step.
- Supports human-in-the-loop review before patching.

2. Conversion Layer (`resume_patcher.py`):
- Applies deterministic document conversion updates to the `.docx`.
- Preserves paragraph/run formatting and layout style.
- Outputs a final tailored `.docx`.

This separation keeps generation flexible and context-aware while keeping document transformation repeatable.

## Optional Resume Preferences

`resume_preferences.json` is optional.

- If present, the chatbot should use it by default when generating tailored content.
- To skip it, simply say: "Ignore `resume_preferences.json` for this run."
- The file guides chatbot content decisions before `replacements.json` is created.
- `resume_preferences.json` is not consumed by `resume_patcher.py`.
- No preference modes or config file are required.

## What `resume_patcher.py` Does

`resume_patcher.py`:

- Opens a source `.docx` resume.
- Applies deterministic text replacements/updates from prepared patch inputs.
- Preserves style by cloning paragraph/run formatting from template paragraphs in the source document.
- Handles multi-line paragraph formatting (for header/title + date/location line breaks).
- Writes the result to a new output `.docx` path.

It should remain a conversion tool only. It should not enforce resume writing rules, tone, strategy, bullet-count policy, or ATS/recruiter optimization logic.

## Requirements

- Python 3.10+
- Microsoft Word-compatible `.docx` source resume

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python3 resume_patcher.py --src "/path/to/master_resume.docx" --replacements "/path/to/replacements.json" --out "/path/to/tailored_resume.docx"
```

Arguments:

- `--src`: Input resume `.docx`
- `--replacements`: Replacement JSON path
- `--out`: Output resume `.docx`

The script prints `Wrote <output_path>` on success.

## Replacement JSON Styles

Block replacements may provide paragraphs as strings or objects. Object items support these deterministic formatting styles:

- `section_heading`: top-level section label template, such as `SUMMARY`, `SKILLS`, `Work EXPERIENCE`, or `EDUCATION`.
- `company_heading`: company/group heading template discovered from existing resume company headings.
- `role_heading`: role/title heading template discovered from existing role headings.
- `date_location`: date/location line template discovered from existing date/location paragraphs.
- `normal`: safe plain paragraph template. This is intentionally never allowed to fall back to Heading 1.
- `bullet`: bullet/list paragraph template.
- `keep`: preserve the target paragraph's current formatting when replacing that paragraph.

Example:

```json
{
  "block_replacements": [
    {
      "start_heading": "Work EXPERIENCE",
      "end_heading": "EDUCATION",
      "content": [
        {"text": "Work EXPERIENCE", "style": "section_heading"},
        {"text": "Pearl Marketing & Freelance | Backend Systems, Automation & Web Application Support", "style": "company_heading"},
        {"text": "Application Support Consultant", "style": "role_heading"},
        {"text": "Remote | May 2025 - Present", "style": "date_location"},
        {"text": "Supported live WordPress and WooCommerce systems...", "style": "bullet"}
      ]
    }
  ]
}
```

The legacy key `replacement_paragraphs` is still supported; `content` is accepted as an alias for block replacement items. Existing files that use only `normal` and `bullet` continue to work, but `normal` now resolves to a safe non-heading template instead of cloning the replaced block's first paragraph.

## End-to-End AI Workflow

1. Provide the chatbot with:
- Target job description
- Source resume `.docx`
- Optional `resume_preferences.json`

2. Ask the chatbot to:
- Propose tailored substitutions aligned with your request
- Keep edits truthful and role-relevant
- Return a reviewed patch plan and `replacements.json`

3. Run `resume_patcher.py` to generate the tailored output `.docx`.

4. Review the final document in Word for content and formatting quality.

## AI Prompt Examples

- "Use the uploaded resume patcher ZIP, but ignore `resume_preferences.json` for this run."
- "Use the uploaded files and create the tailored resume for this JD."

## Important Behavior Notes

- The patch logic is currently profile-specific and uses anchor text lookups (exact/contains matching) for target sections.
- If expected anchor text is missing or changed in the source resume, the script raises a `ValueError`.
- The script discovers semantic paragraph templates from the source document before applying block replacements. It prefers real section, company, role, date/location, normal, and bullet paragraphs already present in the resume.
- The script preserves existing section layout intent and avoids rebuilding the document from scratch.

## Regression Check

Run the focused Work Experience formatting regression after installing dependencies:

```bash
python3 tests/regression_work_experience_styles.py
```

The check replaces `Work EXPERIENCE` through `EDUCATION`, then inspects the generated DOCX XML to verify that inserted company, role, and date/location paragraphs are not Heading 1, the retained section heading is Heading 1, and inserted bullets keep bullet/list formatting.

## Troubleshooting

### `ModuleNotFoundError: No module named 'docx'`

Install dependencies first:

```bash
pip install -r requirements.txt
```

### Anchor text not found

If the source resume text differs from what the script expects, update anchor strings in `resume_patcher.py` or restore the expected source wording.

### Formatting drift after source resume edits

If the source resume structure changed significantly, re-check template paragraph indices and anchor targets in `resume_patcher.py`.
