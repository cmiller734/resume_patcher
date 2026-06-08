# Chatbot Preference Policy

This policy defines how the AI/chatbot layer should decide whether to use `resume_preferences.json` during normal conversation.

## Goal

Keep chatbot UX conversational and low-friction. Users should not need to remember special commands in most sessions.

## Scope

- Applies to the AI/chatbot generation layer only.
- Does not apply to `resume_patcher.py`.
- `resume_patcher.py` remains deterministic document conversion only.

## Preference Usage Rules

1. If the user explicitly says to ignore preferences, do not use `resume_preferences.json` for that run.
2. If the user does not mention preferences and `resume_preferences.json` exists, use it by default to guide generation.
3. If `resume_preferences.json` does not exist, continue without it.

## Intent Mapping Examples

- "don't use resume preferences" -> ignore preferences for this run
- "ignore the preferences file" -> ignore preferences for this run
- "use the preferences file" -> use preferences
- No preference mention -> use preferences when present

## Precedence

1. Factual accuracy is mandatory.
2. Explicit live user instructions override preference guidance.

## Missing/Invalid Preferences File

- If `resume_preferences.json` is missing, unreadable, or invalid JSON:
1. Continue without preferences.
2. Briefly notify the user that preferences were unavailable.

## Deliverables

1. Deliver a resume, always.
2. Deliver a cover letter if the user requests it.
3. No need to provide generated JSON resources