"""Validate that `--workers N` produces equivalent results to single-worker mode.

Runs the solver in single-worker mode (sequential baseline), then in parallel,
and diffs per-project NPP / Dev Fee / FMV / DSCR / Live IRR / Appraisal IRR /
Equity %. Exit code 0 if all fields are within tolerance, 1 otherwise.

Usage:
    python validate_parallel.py <workbook.xlsm>
    python validate_parallel.py <workbook.xlsm> --workers 4 --tolerance 1e-3
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from dn38_solver.validation.parallel_correctness import (
    format_report,
    validate_parallel,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="validate_parallel",
        description=(
            "Parallel-vs-sequential correctness gate for the dn38-solver. "
            "Runs both modes against the same workbook and confirms per-"
            "project outputs agree within tolerance."
        ),
    )
    parser.add_argument("workbook", help="Path to .xlsm workbook to validate against")
    parser.add_argument("--workers", type=int, default=2,
                        help="Number of parallel workers to compare (default: 2)")
    parser.add_argument("--tolerance", type=float, default=1e-4,
                        help="Per-field absolute tolerance (default: 1e-4)")
    parser.add_argument("--timeout", type=int, default=3600,
                        help="Per-pass timeout in seconds (default: 3600)")
    parser.add_argument("--strict", action="store_true",
                        help="Use strict convergence (no relaxed-tier) for both passes")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
        stream=sys.stdout,
    )

    wb_path = Path(args.workbook)
    if not wb_path.exists():
        print(f"ERROR: workbook not found: {wb_path}", file=sys.stderr)
        return 2

    report = validate_parallel(
        wb_path,
        workers=args.workers,
        tolerance=args.tolerance,
        timeout_sec=args.timeout,
        allow_relaxed=not args.strict,
    )
    print(format_report(report))
    return 0 if report.all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
