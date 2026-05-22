# AI Resume Tailoring App

An AI-assisted resume tailoring workflow that combines chatbot judgment with deterministic DOCX patching.

This repository is designed to be used with an AI chatbot. The chatbot analyzes a target job description and (optionally) generates substitution/replacement content using the user preference profile in `resume_preferences.json`. Then `resume_patcher.py` performs the document conversion step by applying updates to a source `.docx` while preserving formatting patterns.

## Overview

This is not designed to be a "resume generation pipeline". The goal of this open-source project is to utilize the latest breakthroughs in technology to allow real people to succeed in their careers. It also helps recruiters and hiring managers by giving them a very well-structured document with which to evaluate their candidates.

This is also not designed to cheat the system. If you make up experience on your resume or claim to have experience you don't actually have, you will be wasting both your time and your recruiter's time. It IS designed to help you present your unique, human, hard-won experience in a way that is helpful for both you AND hiring managers.

## How To Use

1. Clone the repository to your system.
2. Add your master resume to the repository. Include EVERYTHING you've ever done professionally in your field in that master resume.
3. Zip the repository and upload it to the AI chatbot of your choice.
4. Paste the job description into the chat prompt or upload it as a separate txt file (either works).
5. Now that the chatbot has the ZIP and the JD, ask the chatbot to use the attached ZIP and enclosed instructions to create your tailored resume
    NOTE: there's a "resume_preferences.json" file. The chatbot will use that by default, unless you ask it not to.
6. This is where the "process" ends. You will need to work with the finished resume to perform final edits alongside the chatbot. This is also why you can't use this tool to "cheat the system"; you will still need to do manual work for your resume to impress an actual human.

## Repository Contents

- `resume_patcher.py`: Deterministic DOCX conversion/patching engine.
- `resume_preferences.json` (optional): User preference profile for AI generation decisions (tone, bullet density, framing, ATS balance, credibility guardrails).
- `resume_generation_config.example.json` (optional example): AI-layer config pattern showing preference mode selection.
- `requirements.txt`: Python dependency list for the patcher.

## App Architecture

This workflow intentionally separates responsibilities:

1. AI Layer (chatbot):
- Reads the job description and source resume context.
- Optionally uses `resume_preferences.json` to guide tailoring choices.
- Produces proposed substitutions/revisions (optionally as structured JSON for review).
- Creates `replacements.json` before any patching step.
- Supports human-in-the-loop review before patching.

2. Conversion Layer (`resume_patcher.py`):
- Applies deterministic document conversion updates to the `.docx`.
- Preserves paragraph/run formatting and layout style.
- Outputs a final tailored `.docx`.

This separation keeps generation flexible and context-aware while keeping document transformation repeatable.

## How `resume_preferences.json` Is Used

`resume_preferences.json` is optional and is not consumed by `resume_patcher.py`.

Instead, it is used by the AI layer to shape content generation, including:

- Voice/tone constraints (specific, credible, senior, non-generic).
- Content priorities (implementation ownership, troubleshooting depth, customer/business impact).
- Bullet density and relevance strategy.
- Job-title framing guidance.
- ATS keyword balance with human readability.

In practice: when present, the chatbot should read this file before generating replacements/substitutions and before creating `replacements.json`.

## Preference Modes

Preference mode applies to the AI/chatbot generation layer only:

- `none`: Ignore/omit preferences. Use only the master resume, job description, live user instructions, and general resume best practices.
- `guided`: Use preferences as soft guidance. This is the recommended default.
- `strict`: Treat preferences as hard constraints where possible, while preserving factual accuracy and explicit user instructions.

If preferences are missing, the workflow still runs using `none` behavior.

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
python3 resume_patcher.py --src "/path/to/master_resume.docx" --out "/path/to/tailored_resume.docx"
```

Arguments:

- `--src`: Input resume `.docx`
- `--out`: Output resume `.docx`

The script prints `Wrote <output_path>` on success.

## End-to-End AI Workflow

1. Provide the chatbot with:
- Target job description
- Source resume `.docx`
- Optional `resume_preferences.json`

2. Ask the chatbot to:
- Propose tailored substitutions aligned with preferences
- Keep edits truthful and role-relevant
- Return a reviewed patch plan and `replacements.json`

3. Run `resume_patcher.py` to generate the tailored output `.docx`.

4. Review the final document in Word for content and formatting quality.

## AI Prompt Examples

- "Use the uploaded resume patcher ZIP, but ignore `resume_preferences.json` for this run."
- "Use `resume_preferences.json` as guided preferences, not hard constraints."
- "Use `resume_preferences.json` in strict mode. Flag conflicts instead of silently making bad edits."

## Important Behavior Notes

- The patch logic is currently profile-specific and uses anchor text lookups (exact/contains matching) for target sections.
- If expected anchor text is missing or changed in the source resume, the script raises a `ValueError`.
- The script depends on template paragraph indices from the source document for style cloning (currently paragraphs `8` and `10`). Large structural changes to the master resume may require script updates.
- The script preserves existing section layout intent and avoids rebuilding the document from scratch.

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
