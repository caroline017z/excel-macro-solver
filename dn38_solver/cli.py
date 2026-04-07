"""dn38_solver.cli — Command-line interface.

Usage:
    python -m dn38_solver.cli                           # Solve with defaults
    python -m dn38_solver.cli path/to/workbook.xlsm     # Specific workbook
    python -m dn38_solver.cli --dry-run                  # Inspect without solving
    python -m dn38_solver.cli --history                  # Show recent runs
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from dn38_solver.config import DEFAULT_WORKBOOK, OUTPUT_ROWS
from dn38_solver.solver.orchestrator import solve_all
from dn38_solver.storage.database import get_connection, get_runs


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(message)s",
        stream=sys.stdout,
    )


def _show_history(limit: int = 20) -> None:
    conn = get_connection()
    runs = get_runs(conn, limit)
    conn.close()

    if not runs:
        print("No runs recorded yet.")
        return

    print(f"\n{'='*80}")
    print(f"  Recent Solver Runs")
    print(f"{'='*80}")
    print(f"  {'Timestamp':<22} {'Workbook':<30} {'Projects':>8} {'Status':>14}")
    print(f"  {'-'*22} {'-'*30} {'-'*8} {'-'*14}")

    for r in runs:
        ts = r.run_timestamp[:19]
        wb = r.workbook_name[:28]
        n = len(r.projects)
        print(f"  {ts:<22} {wb:<30} {n:>8} {r.status:>14}")

        for p in r.projects:
            npp = f"${p.npp_per_w:.3f}" if p.npp_per_w else "—"
            dev = f"${p.dev_fee_per_w:.3f}" if p.dev_fee_per_w else "—"
            stat = "OK" if p.converged else "CHECK"
            print(f"    {p.name:<26} NPP={npp:>8}  DevFee={dev:>8}  [{stat}]")


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="dn38-solver",
        description="38DN Hybrid Shadow Solver — No Excel COM in your session",
    )
    parser.add_argument(
        "workbook",
        nargs="?",
        default=None,
        help=f"Path to .xlsm workbook (default: {DEFAULT_WORKBOOK.name})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Read workbook and list projects without solving",
    )
    parser.add_argument(
        "--history",
        action="store_true",
        help="Show recent solver run history",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=600,
        help="COM subprocess timeout in seconds (default: 600)",
    )
    parser.add_argument(
        "--batch-id",
        default=None,
        help="Custom batch ID for grouping runs",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable debug logging",
    )

    args = parser.parse_args()
    _setup_logging(args.verbose)

    if args.history:
        _show_history()
        return

    workbook_path = Path(args.workbook) if args.workbook else DEFAULT_WORKBOOK

    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        sys.exit(1)

    record = solve_all(
        workbook_path,
        batch_id=args.batch_id,
        dry_run=args.dry_run,
        timeout_sec=args.timeout,
    )

    # Exit code: 0 if converged, 1 if not
    match record.status:
        case "converged" | "dry_run":
            sys.exit(0)
        case _:
            sys.exit(1)


if __name__ == "__main__":
    main()
