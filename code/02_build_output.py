"""Entrypoint: build the final output Excel from inputs.

Reads:
    - input/indicators.xlsx (codebook)
    - input/scores.xlsx (scores + analysis from Claude Code)
    - input/states.yaml (state list)

Writes:
    - output/State_Constitutional_Analysis.xlsx

Usage:
    python code/02_build_output.py
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

# Allow running directly without installing
sys.path.insert(0, str(Path(__file__).parent))

from lib.excel_builder import (
    build_workbook,
    load_indicators,
    load_scores,
    load_states,
)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        default=None,
        help="Output file path (default: output/State_Constitutional_Analysis.xlsx)",
    )
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    root = Path(__file__).resolve().parent.parent

    indicators_path = root / "input" / "indicators.xlsx"
    scores_path = root / "input" / "scores.xlsx"
    states_path = root / "input" / "states.yaml"

    if not indicators_path.exists():
        print(f"ERROR: {indicators_path} not found. Run code/00_build_codebook.py first.")
        sys.exit(1)
    if not states_path.exists():
        print(f"ERROR: {states_path} not found.")
        sys.exit(1)

    indicators = load_indicators(indicators_path)
    states = load_states(states_path)
    scores, analysis = load_scores(scores_path)

    print(
        f"Loaded {len(indicators)} indicators, {len(states)} states, "
        f"{len(scores)} scores, {len(analysis)} analysis sections"
    )

    if not scores:
        print(
            "\nWARNING: input/scores.xlsx not found or empty. The output will have\n"
            "blank score cells. Run Claude Code's /analyze command first.\n"
        )

    output_path = (
        Path(args.output)
        if args.output
        else root / "output" / "State_Constitutional_Analysis.xlsx"
    )

    build_workbook(indicators, states, scores, analysis, output_path)
    print(f"Done: {output_path}")


if __name__ == "__main__":
    main()
