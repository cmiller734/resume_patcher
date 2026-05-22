# Resume Tailoring Toolkit

A lightweight workflow for tailoring a master resume to specific job descriptions while preserving the original `.docx` formatting.

This toolkit is designed for repeatable resume tailoring: the AI assistant generates a targeted replacements JSON file, then the patching script applies those replacements directly into the master Word document. The goal is to keep the resume visually consistent while making focused, role-specific edits for ATS alignment and human readability.

## Purpose

This process exists to avoid rebuilding resumes from scratch.

Instead of creating a new document every time, the workflow uses a trusted master resume as the formatting and content source. The assistant analyzes a job description, decides which sections should be adjusted, creates a replacements JSON object, and the script patches the existing `.docx` file.

The result should feel like the original resume, but sharper and more targeted for the role.

## Toolkit Files

Recommended package structure:

```text
resume_tailoring_toolkit/
   Master Resume.docx
  resume_patch_docx_workhorse_linebreaks.py
  README.md
```

### ` Master Resume.docx`

The source resume.

This file should be treated as the master template for both formatting and content. The patching process should preserve its layout, section order, spacing, bullet style, font hierarchy, and general visual feel.

Do not rebuild the resume from scratch unless intentionally redesigning it.

### `resume_patch_docx_workhorse_linebreaks.py`

The patching script.

This script applies exact text replacements to the master `.docx` while preserving Word formatting as much as possible. It is intended to be used with a JSON object containing `find` and `replace` pairs.

### `README.md`

This file.

It explains how to package, upload, and run the resume tailoring workflow consistently.

## Workflow Overview

1. Start with the master resume `.docx`.
2. Provide a target job description.
3. Generate a replacements JSON object tailored to the role.
4. Validate that each `find` string exists in the source resume.
5. Run the patching script.
6. Review the generated `.docx` for formatting and content quality.
7. Export or submit the tailored resume.

## AI Assistant Instructions

When using this toolkit with an AI assistant, provide the job description and upload the toolkit files.

Use a prompt like:

```text
Use the uploaded master resume and patcher script.
Tailor the resume for this job description.
Preserve the original document formatting and structure.
Generate the replacements JSON, run the patcher, and return the finished `.docx`.
Also summarize the main changes.
```

The assistant should:

- Use ` Master Resume.docx` as the source of truth.
- Preserve the original resume formatting.
- Make targeted copy edits rather than redesigning the document.
- Keep bullets specific, credible, and grounded in actual experience.
- Avoid generic AI-sounding resume language.
- Prioritize ATS alignment without flattening the resume into keyword soup.
- Emphasize the experience most relevant to the target role.
- Return the finished `.docx` file with an employer-friendly filename.

## Replacement JSON Format

The patcher expects a JSON structure containing exact text replacements.

Example:

```json
{
  "replacements": [
    {
      "find": "Technical Support Engineer with 9 years of experience supporting enterprise SaaS platforms, customer-facing application environments, and production business systems.",
      "replace": "Technical Customer Success Engineer with 9 years of experience supporting enterprise SaaS platforms, customer-facing application environments, and production business systems."
    }
  ]
}
```

Each `find` value should match text that already exists in the master resume.

Each `replace` value should be the revised text that should appear in the tailored resume.

## Important Rules

### Preserve Formatting

The master resume is the formatting template.

Do not intentionally change:

- Layout
- Section order
- Bullet style
- Font hierarchy
- Spacing system
- Overall visual structure

The goal is to patch text into the existing document, not create a new resume from scratch.

### Keep Replacements Exact

The `find` strings must match the current resume text exactly.

If a `find` string does not match, the patch may fail or silently skip that replacement. Before running the patcher, validate that every `find` string exists in the source document.

### Prefer Focused Edits

Most tailoring should happen in:

- Summary
- Skills
- Most recent experience
- Apica Professional Services Engineer section
- Apica Senior Technical Support Engineer section
- Selected project bullets when relevant

Avoid rewriting every role unless the job description truly calls for it.

### Keep the Resume Honest

Do not invent experience.

Good tailoring means reframing real experience toward the target role, not creating false qualifications.

For example:

- Strong: “troubleshot network, API, and agent connectivity issues across production environments”
- Weak/false: “administered 5G mobile core infrastructure” unless that experience actually exists

## Recommended Tailoring Strategy

For technical support, customer success engineering, implementation, SaaS, and application support roles, emphasize:

- Enterprise SaaS support
- Technical onboarding
- Escalation ownership
- API troubleshooting
- Logs, SQL, telemetry, and Postman
- Windows/Linux environments
- Customer communication
- Documentation and runbooks
- Engineering collaboration
- Root cause analysis
- Production troubleshooting
- Implementation work

For roles involving cloud infrastructure, networking, IoT, cellular, or cybersecurity, emphasize adjacent experience honestly:

- Secure enterprise environments
- Network troubleshooting
- SSL/TLS certificates
- DNS
- Monitoring agents
- On-premise deployments
- Production connectivity issues
- Authentication and access management
- Customer-owned infrastructure
- Incident response and outage coordination

## Suggested Output Naming

Use employer-friendly lowercase filenames:

```text
caleb_miller_monogoto_resume.docx
caleb_miller_companyname_resume.docx
caleb_miller_technical_customer_success_engineer_resume.docx
```

Avoid spaces when possible.

## Quality Checklist

Before using the tailored resume, verify:

- The file opens correctly in Microsoft Word.
- Header/contact info is intact.
- Summary is aligned with the target role.
- Skills section contains relevant keywords from the job description.
- Bullets remain specific and believable.
- No experience has been exaggerated beyond what is accurate.
- Formatting is consistent with the master resume.
- No section was accidentally removed.
- Education is still present.
- The final filename is clean and employer-friendly.

## Common Failure Modes

### Missing Script

If the patcher script is not available in the current session, upload it again or include it in the toolkit zip.

### Missing Master Resume

If the master resume is not available in the current session, upload it again or include it in the toolkit zip.

### Replacement Text Does Not Match

If a `find` string fails, copy the exact text from the current master resume and update the JSON.

Common causes:

- Smart quotes vs. straight quotes
- Different dash characters
- Extra spaces
- Line breaks
- Bullet characters
- Recently edited resume text

### Formatting Looks Wrong

This usually means the resume was rebuilt instead of patched, or the replacement text introduced awkward line breaks.

Use smaller, more targeted replacements and preserve the original structure.

## Recommended Session Starter

At the start of a resume tailoring session, upload either:

```text
resume_tailoring_toolkit.zip
```

or the two required files:

```text
 Master Resume.docx
resume_patch_docx_workhorse_linebreaks.py
```

Then provide the job description and ask the assistant to run the tailoring process.

## Philosophy

The resume should not sound like a generic AI rewrite.

It should sound like Caleb: technically credible, specific, practical, and grounded in real support, implementation, troubleshooting, and customer-facing engineering work.

The point is not to become a different candidate for every job.

The point is to help the right parts of the existing experience show up clearly for the specific role.

