---
description: Run the de jure constitutional analysis for all states
allowed-tools: Read, Write, Edit, Bash(python:*), Bash(ls:*), Bash(cat:*), Bash(find:*), Bash(wc:*), Bash(head:*)
---

# Constitutional Analysis Task

You are conducting a **de jure comparative analysis of U.S. state constitutions** to support pilot state selection for a research study.

## Inputs to read

1. **The full instructions** in `code/prompts/evaluation_prompt.md` — read this first and follow it carefully.
2. **The codebook** in `input/indicators.xlsx` — this defines the indicators, questions, and scoring guidance.
3. **The constitutions** in `data/raw/*.txt` — one file per state (south_carolina.txt, tennessee.txt, texas.txt, utah.txt, minnesota.txt, ohio.txt).
4. **The states config** in `input/states.yaml` — state codes and metadata.

## Your output

Produce `input/scores.xlsx` with two sheets:

### Sheet `scores`
Tidy long-format table, one row per (indicator × state) combination, with columns:
- `indicator_id` (e.g., FF01)
- `state_code` (SC, TN, TX, UT, MN, OH)
- `score` (integer 0, 1, 2, or 3 — follow the scoring_guidance in the codebook strictly)
- `constitutional_reference` (e.g., "Article I §3" or "Not provided" if absent)
- `note` (1–3 sentence justification grounded in the constitutional text)

### Sheet `analysis`
Narrative analysis with columns `section_id` and `content`. Required sections:
- `executive_summary` — 2–3 paragraph overview of findings
- `key_pattern` — the most important comparative pattern observed
- `profile_SC`, `profile_TN`, `profile_TX`, `profile_UT`, `profile_MN`, `profile_OH` — one paragraph each on the state's distinctive constitutional features
- `selection_guidance` — recommendation on pilot selection logic given the findings

## Critical guidelines

1. **Read `code/prompts/evaluation_prompt.md` in full before starting.** That document is authoritative.
2. **Ground every score in actual constitutional text.** Quote or cite the article/section where possible.
3. **De jure only.** Do not score based on practice, jurisprudence, or general state reputation. Score what the text says.
4. **Equal weighting.** Do not skip indicators; produce one row per (indicator × state). If a provision is absent, that is itself a score (usually 0).
5. **Use openpyxl** to write `input/scores.xlsx`. Use Arial font. No need for elaborate formatting — `02_build_output.py` produces the styled final output.

## Starting steps

1. Read `code/prompts/evaluation_prompt.md`
2. Verify all 6 constitution files exist in `data/raw/`. If any are missing, alert the user and stop.
3. Read `input/indicators.xlsx` to get the indicator list.
4. For each state, read the full constitution text.
5. For each (indicator × state), assess and score.
6. Write `input/scores.xlsx` with both sheets.
7. Confirm completion with a one-line summary of total rows written.

Begin.
