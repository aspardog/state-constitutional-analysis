# State Constitutional Analysis

A reproducible pipeline for de jure comparative analysis of U.S. state constitutions, designed to support pilot state selection for studies on **Fundamental Freedoms**, **Checks and Balances**, and **Democratic Protections**.

This project is designed to run with **[Claude Code](https://docs.claude.com/en/docs/claude-code)** for the LLM-based evaluation step. The Claude Code project files needed for that workflow are included under `.claude/`.

No Excel files are committed to this repository. Any `.xlsx` inputs or outputs are expected to remain local only.

## States Covered

- South Carolina (SC)
- Tennessee (TN)
- Texas (TX)
- Utah (UT)
- Minnesota (MN)
- Ohio (OH)

## Architecture

The full research workflow has three stages:

```
┌─────────────────┐      ┌──────────────────────┐      ┌──────────────────┐
│ Private scrape  │ ───> │  Claude Code evaluates │ ───> │ 02_build_output  │
│  (optional)     │      │  (one-time, manual)    │      │  (run anytime)   │
└─────────────────┘      └──────────────────────┘      └──────────────────┘
        │                          │                            │
        v                          v                            v
   data/raw/*.txt          input/scores.xlsx         local .xlsx outputs
   (cached locally)        (canonical local input)   (not committed)
```

### Why this design?

- **Constitution text is treated as an input artifact.** You can generate it with your own scraper or add it manually to `data/raw/`.
- **LLM evaluation is non-deterministic but stable as a frozen artifact.** Claude Code runs the evaluation once, producing `input/scores.xlsx`. From that moment on, `scores.xlsx` is the canonical local source. You can edit it by hand if you disagree with any score.
- **Output building is deterministic.** Re-run `02_build_output.py` as many times as you want; it always produces the same Excel from the same inputs.

## Folder Structure

```
state-constitutional-analysis/
├── .claude/                        Claude Code project config
│   ├── settings.json
│   └── commands/
│       └── analyze.md              Custom slash command: /analyze
├── code/
│   ├── 02_build_output.py          Output builder entrypoint
│   ├── lib/                        Reusable modules
│   │   └── excel_builder.py
│   ├── prompts/
│   │   └── evaluation_prompt.md    The prompt Claude Code follows
│   └── tests/
│       └── test_excel_builder.py
├── input/
│   ├── indicators.xlsx             Local editable codebook (gitignored)
│   ├── states.yaml                 Six states + scraping config
│   └── scores.xlsx                 Generated locally by Claude Code (gitignored)
├── data/
│   └── raw/                        Constitution text inputs (gitignored)
└── outputs/                        Optional local export location (gitignored)
```

## Setup

```bash
# Clone and enter
git clone <your-repo-url>
cd state-constitutional-analysis

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Step 1 — Populate `data/raw/` with constitution text

```bash
mkdir -p data/raw
```

Add one plain-text file per state to `data/raw/` before running the analysis:

- `south_carolina.txt`
- `tennessee.txt`
- `texas.txt`
- `utah.txt`
- `minnesota.txt`
- `ohio.txt`

You can create these files with your own scraper or by manual copy-paste. They are intentionally gitignored.

### Step 2 — Run the LLM evaluation with Claude Code (one-time)

Open this project in Claude Code and run:

```
/analyze
```

This triggers the custom command defined in `.claude/commands/analyze.md`. Claude Code will:

1. Read the codebook from `input/indicators.xlsx`
2. Read each constitution from `data/raw/`
3. For each (indicator × state), assign a score 0–3, a constitutional reference, and a brief note
4. Write the local `input/scores.xlsx` with two sheets: `scores` and `analysis`

This step takes several minutes and is the most expensive part of the pipeline (LLM context tokens). Once `scores.xlsx` exists, treat it as the canonical local input — edit by hand if you disagree with any score.

### Step 3 — Build the output Excel

```bash
python code/02_build_output.py
```

This produces a local `.xlsx` workbook with the following sheets:

1. **Read Me First** — purpose, glossary, caveats
2. **Executive Summary** — aggregate scores by dimension and total
3. **FF — Comparative** — Fundamental Freedoms, indicator-by-indicator
4. **CB — Comparative** — Checks and Balances
5. **DP — Comparative** — Democratic Protections
6. **Tidy Data** — long format for analysis (one row per indicator × state)
7. **Analysis** — narrative analysis from Claude Code's evaluation
8. **Sources** — references with hyperlinks

You can re-run `02_build_output.py` after any manual edit to `input/scores.xlsx` or `input/indicators.xlsx`.

The generated workbook should remain local; it is gitignored by default.

## Editing the codebook

Your local `input/indicators.xlsx` is the source of truth for indicators. Columns:

| Column | Description |
|---|---|
| `id` | Stable identifier (e.g., FF01, CB03, DP07) |
| `dimension` | One of: `Fundamental Freedoms`, `Checks and Balances`, `Democratic Protections` |
| `indicator` | Short name of the indicator |
| `question` | The question Claude Code answers about each constitution |
| `scoring_guidance` | Explicit definition of what 0/1/2/3 means for this indicator |

To add an indicator, just add a row. To remove one, delete the row. To re-score, delete `input/scores.xlsx` and re-run `/analyze` to regenerate all scores.

## Methodology Notes

- All analysis is **de jure** (textual only). No statutory law, case law, or implementation practice is captured.
- Scoring is **ordinal 0–3** with direction defined per indicator. Higher is more democratic / more protective / more checks.
- Aggregate scores are **simple sums** with equal weighting. To use weighted scoring, derive new aggregates from the `Tidy Data` sheet.
- See the `Read Me First` sheet in the output for the full methodology.

## Running tests

```bash
pytest code/tests/ -v
```

Tests cover the Excel builder output structure and formula integrity.

## License

MIT — see `LICENSE`.

## Citation

If you use this in published work, please cite as:

```
State Constitutional Analysis (2026). De jure comparative of SC, TN, TX, UT, MN, OH.
Internal research tool. https://github.com/<your-org>/state-constitutional-analysis
```
