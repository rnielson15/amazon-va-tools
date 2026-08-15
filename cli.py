#!/usr/bin/env python3
"""
Command-line entry point.

    python cli.py INPUT_BULK.xlsx -o OUTPUT.xlsx
    python cli.py INPUT_BULK.xlsx                 # -> Bleeders_Report.xlsx
"""

import argparse
import sys

from bleeders import generate, save_workbook


def main():
    ap = argparse.ArgumentParser(description="Build an Amazon Bleeders Report from a bulk export.")
    ap.add_argument("input", help="Path to the Amazon bulk operations .xlsx")
    ap.add_argument("-o", "--output", default="Bleeders_Report.xlsx")
    args = ap.parse_args()

    try:
        summary, wb = generate(args.input)
    except FileNotFoundError:
        sys.exit(f"Input file not found: {args.input}")

    save_workbook(wb, args.output)

    total = sum(n for _, n in summary)
    print("Bleeders Report generated:", args.output)
    print("-" * 44)
    for name, n in summary:
        print(f"  {name:<34} {n:>4}")
    print("-" * 44)
    print(f"  {'TOTAL bleeders':<34} {total:>4}")
    if total == 0:
        print("\nNo bleeders found (or the relevant source tabs were empty).")


if __name__ == "__main__":
    main()
