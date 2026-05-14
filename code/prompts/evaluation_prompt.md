# Constitutional Analysis Evaluation Prompt

You are conducting a **de jure comparative analysis** of U.S. state constitutions to support pilot state selection for a research study on Fundamental Freedoms, Checks and Balances, and Democratic Protections.

This document is your authoritative instruction set. Read it carefully and follow it precisely.

---

## Your task

For each combination of (indicator × state), produce:
- A score (integer 0, 1, 2, or 3)
- A constitutional reference (e.g., "Article I §3")
- A 1–3 sentence note justifying the score

Then, write narrative analysis sections summarizing patterns and per-state findings.

Output everything to `input/scores.xlsx` with two sheets: `scores` and `analysis`.

---

## Inputs to read

Read these files in this order:

1. **Codebook**: `input/indicators.xlsx` (sheet `indicators`)
   - Columns: `id`, `dimension`, `indicator`, `question`, `scoring_guidance`
   - This defines all the questions you must answer for each state
2. **States config**: `input/states.yaml`
   - Defines the 6 state codes and names
3. **Constitution texts**: `data/raw/{state_filename}.txt`
   - One file per state. Filenames are the state name lowercased with underscores (e.g., `south_carolina.txt`, `tennessee.txt`)
   - If any file is missing or under ~1000 words, STOP and alert the user. Do not invent content.

---

## Core methodological principles

### Principle 1 — De jure only

Score what the **constitutional text** says, not what practice does. Specifically:

- DO NOT use your general knowledge of state politics, court rulings, or statutory law
- DO NOT score based on reputation (e.g., "Texas is known for X")
- DO use the explicit text of the constitution and its amendments
- IF a clause is in the text but is known to be unenforceable (e.g., religious tests under Torcaso v. Watkins), score the TEXT, and note the unenforceability in the note field
- IF the text is silent on something, that absence is itself a finding (usually a low score)

### Principle 2 — Grounding in text

Every score must be grounded in identifiable text:
- For non-zero scores, you MUST cite the Article and Section (e.g., "Article I §3")
- If the score is 0 because the provision is absent, write "Not provided" in the reference field
- The note field must explain WHY the score is what it is, with a brief quote or paraphrase

### Principle 3 — Strict adherence to scoring guidance

The `scoring_guidance` column in the codebook defines what 0/1/2/3 mean for each indicator. Apply it literally.

If you find yourself rationalizing a score that doesn't match the guidance, you are doing it wrong. Either:
- The score matches the guidance → use it
- The guidance doesn't fit the constitutional reality → use the closest match and explain in the note

### Principle 4 — Be consistent across states

Apply the same standard to each state. If you give Utah a 3 for "religious freedom and church-state separation" because the text has X and Y characteristics, the same X and Y characteristics in Ohio should also be a 3.

### Principle 5 — Complete coverage

Produce **every (indicator × state) combination** as a row. If there are 22 indicators and 6 states, you should have 132 rows in the `scores` sheet. Do not skip any combination, even if a provision is absent (use score 0).

---

## How to evaluate each indicator

For each (indicator, state) pair:

1. **Read the indicator definition** carefully — the question and scoring_guidance
2. **Search the constitution text** for relevant provisions:
   - Use keyword search (grep, ctrl-F mentally) for key terms: "religion", "judge", "amend", "vote", "rights", "bail", "speech", etc.
   - Also check by structure: Article I (rights), Article on judiciary, Article on amendments, Article on suffrage
3. **Read the relevant clauses in context** — don't take a clause out of context
4. **Apply the scoring guidance** to the text you found
5. **Write the cite and the note**

For complex indicators (e.g., plural executive composition, judicial selection method), you may need to read multiple articles. Be thorough.

---

## Output specification

Write `input/scores.xlsx` with these two sheets.

### Sheet `scores` — tidy long format

| Column | Type | Description |
|---|---|---|
| `indicator_id` | text | Exact match to id column in codebook (e.g., FF01) |
| `state_code` | text | Two-letter code: SC, TN, TX, UT, MN, OH |
| `score` | int | 0, 1, 2, or 3 |
| `constitutional_reference` | text | E.g., "Article I §3" or "Not provided" |
| `note` | text | 1–3 sentence justification |

Number of rows = number of indicators × number of states (e.g., 22 × 6 = 132).

### Sheet `analysis` — narrative

| Column | Type | Description |
|---|---|---|
| `section_id` | text | Section identifier (see required sections below) |
| `content` | text | The narrative content for that section |

#### Required sections:

| `section_id` | Content guidance |
|---|---|
| `executive_summary` | 2–3 paragraph overview: which state has the highest/lowest aggregate scores, the most striking comparative finding, and any caveats about the de jure analysis |
| `key_pattern` | One paragraph: the most important comparative pattern observed (e.g., "All four originally-targeted states except Utah lack any direct democracy mechanisms; this creates a sharp Utah-vs-rest division" or similar) |
| `profile_SC` | One paragraph: distinctive constitutional features of South Carolina with implications for the pilot study |
| `profile_TN` | One paragraph: same for Tennessee |
| `profile_TX` | One paragraph: same for Texas |
| `profile_UT` | One paragraph: same for Utah |
| `profile_MN` | One paragraph: same for Minnesota |
| `profile_OH` | One paragraph: same for Ohio |
| `selection_guidance` | One paragraph: recommendation on pilot selection logic given the findings. Consider lenses: most-different case, typical case, theoretical sampling, maximum gap |

---

## Implementation steps

Suggested workflow:

```python
# 1. Read the codebook
import pandas as pd
indicators = pd.read_excel('input/indicators.xlsx', sheet_name='indicators')

# 2. Load states config
import yaml
with open('input/states.yaml') as f:
    states = yaml.safe_load(f)['states']

# 3. For each state, read the constitution into a string variable
# 4. For each (indicator, state), determine score/ref/note
# 5. Assemble into a DataFrame and write
from openpyxl import Workbook
wb = Workbook()
# ... write scores sheet
# ... write analysis sheet
wb.save('input/scores.xlsx')
```

But you do not need to write all 132 rows in a single Python call. You can:
- Work through indicators one at a time
- Build up a list of dicts and only write the Excel at the end
- Iterate state-by-state, asking yourself the indicator questions

Pick the strategy that lets you give each (indicator, state) pair careful attention.

---

## Common pitfalls to avoid

1. **DO NOT confuse state and federal constitutional features.** The Equal Protection Clause of the 14th Amendment is FEDERAL. State constitutions have their own equal protection clauses — find those in the state text.

2. **DO NOT score based on what you think the state's politics are.** Texas's reputation for executive power is irrelevant; what matters is what Article 4 of the Texas constitution actually says.

3. **DO NOT skip an indicator because the constitution doesn't mention it.** Absence is a 0 score, not a "skip".

4. **DO NOT give the same boilerplate note for multiple states.** Each note should reference specific text from that state's constitution.

5. **DO NOT exceed 3 sentences in the note field.** Be concise.

6. **DO read the full Article/Section before scoring**, not just the first paragraph. Some constitutions bury qualifiers and exceptions in later sections.

7. **DO check amendments.** Tennessee's §36 on abortion was added in 2014. The text you're reading should include all current amendments.

---

## Quality check before writing the file

Before saving `input/scores.xlsx`, verify:

- [ ] Every (indicator_id, state_code) combination has exactly one row
- [ ] All scores are integers in {0, 1, 2, 3}
- [ ] All `constitutional_reference` fields are populated (use "Not provided" if absent)
- [ ] All `note` fields are populated and specific to that state
- [ ] The `analysis` sheet has all required sections
- [ ] No section content is empty or placeholder

---

## When done

Print a one-line summary:

```
Wrote input/scores.xlsx: N indicators × M states = K score rows; J analysis sections.
```

Then stop. Do not run `02_build_output.py` — let the user do that.
