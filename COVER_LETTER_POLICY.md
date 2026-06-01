# Cover Letter Chatbot Policy

This policy guides the AI/chatbot layer when creating `cover_letter.json`.

## Scope

- Applies only to cover letter writing in the chatbot layer.
- Does not apply to `resume_patcher.py`.
- The patcher renders final text exactly as provided and performs mechanical validation only.

## Core Rules

1. Use the job description, resume context, `cover_letter_preferences.json`, and live user instructions to write a tailored cover letter.
2. Treat live user instructions as run-specific. A detail such as liking a company's location should be used only when the user says it for that job.
3. Do not treat old text in `Caleb Miller Cover Letter.docx` as source content, structure, strategy, or wording guidance.
4. Keep the letter truthful and grounded in Caleb Miller's actual experience.
5. Do not invent motivations, company facts, accomplishments, credentials, employment history, or personal details.
6. Avoid double hyphens (`--`). Use commas, periods, parentheses, or revised sentence structure instead.
7. Resolve placeholders before producing JSON. Do not leave tokens such as `{{COMPANY}}`.

## Output Contract

Produce a separate `cover_letter.json` file:

```json
{
  "cover_letter": {
    "paragraphs": [
      "Dear Hiring Team,",
      "Body paragraph one...",
      "Body paragraph two...",
      "Sincerely\nCaleb Miller"
    ]
  }
}
```

Each paragraph must be a non-empty string. The patcher will not rewrite, shorten, expand, personalize, or repair this text.
