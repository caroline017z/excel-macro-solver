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
from dn38_solver.storage.database import (
    get_checkpointed_projects,
    get_connection,
    get_runs,
)
from dn38_solver.types import (
    ProjectResult,
    RELAXED_LEGEND,
    SolveStatus,
    convergence_label,
)


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(message)s",
        stream=sys.stdout,
    )


def _show_checkpoints(batch_id: str) -> None:
    """Print every per-project checkpoint stored under `batch_id`.

    Use after a chunked solve has crashed mid-portfolio to see which
    projects landed before the failure. Successful runs clear their
    checkpoints automatically, so a non-empty result implies an
    incomplete run worth investigating.
    """
    conn = get_connection()
    try:
        projects = get_checkpointed_projects(conn, batch_id)
    finally:
        conn.close()

    if not projects:
        print(f"No checkpoints for batch_id={batch_id}.")
        return

    print(f"\n{'='*80}")
    print(f"  Checkpoints for batch_id={batch_id}  ({len(projects)} project(s))")
    print(f"{'='*80}")
    print(f"  {'Project':<32} {'Col':>4} {'NPP':>10} {'DevFee':>10} {'DSCR':>9} {'Eq%':>7} {'Conv':>6}")
    print(f"  {'-'*32} {'-'*4} {'-'*10} {'-'*10} {'-'*9} {'-'*7} {'-'*6}")
    has_relaxed = False
    for p in projects:
        npp = f"${p.npp_per_w:.3f}" if p.npp_per_w is not None else "—"
        dev = f"${p.dev_fee_per_w:.3f}" if p.dev_fee_per_w is not None else "—"
        dscr = f"{p.dscr_multiple:.3f}x" if p.dscr_multiple is not None else "—"
        eq = f"{p.equity_pct*100:.2f}%" if p.equity_pct is not None else "—"
        conv = convergence_label(p)
        if conv == "OK*":
            has_relaxed = True
        print(f"  {p.name[:32]:<32} {p.col:>4} {npp:>10} {dev:>10} {dscr:>9} {eq:>7} {conv:>6}")
    if has_relaxed:
        print(f"  {RELAXED_LEGEND}")


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

    has_relaxed = False
    for r in runs:
        ts = r.run_timestamp[:19]
        wb = r.workbook_name[:28]
        n = len(r.projects)
        print(f"  {ts:<22} {wb:<30} {n:>8} {r.status:>14}")

        for p in r.projects:
            npp = f"${p.npp_per_w:.3f}" if p.npp_per_w else "—"
            dev = f"${p.dev_fee_per_w:.3f}" if p.dev_fee_per_w else "—"
            stat = convergence_label(p)
            if stat == "OK*":
                has_relaxed = True
            print(f"    {p.name:<26} NPP={npp:>8}  DevFee={dev:>8}  [{stat}]")

    if has_relaxed:
        print(f"  {RELAXED_LEGEND}")


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
        "--show-checkpoints",
        metavar="BATCH_ID",
        default=None,
        help=(
            "Print per-project checkpoints stored under BATCH_ID. "
            "Used to inspect what landed before a chunked-solve crash. "
            "Successful runs clear their own checkpoints, so a "
            "non-empty result indicates an incomplete prior run."
        ),
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=1800,
        help=(
            "Solver macro timeout threshold in seconds (default: 1800). "
            "Real portfolios run 700-1200s on the chunked path; the threshold "
            "fires the run-level error status post-hoc, not an in-flight cancel."
        ),
    )
    parser.add_argument(
        "--batch-id",
        default=None,
        help="Custom batch ID for grouping runs",
    )
    parser.add_argument(
        "--strict-validation",
        action="store_true",
        help=(
            "Treat formula errors (#REF!, #DIV/0!, etc.) in the input "
            "workbook or the saved _SOLVED.xlsx as a run failure. Default "
            "is to log them as warnings and proceed."
        ),
    )
    parser.add_argument(
        "--chunked",
        action="store_true",
        help=(
            "Run the macro through per-project entry points (Init / "
            "SolveOneProjectByColHL / Finalize) instead of single-shot "
            "SolveHeadless. Each project is its own COM call so total "
            "macro time can exceed Excel's ~900s RPC timeout. "
            "Recommended for portfolios with 10+ projects or any cold "
            "portfolio."
        ),
    )
    parser.add_argument(
        "--allow-relaxed",
        action="store_true",
        help=(
            "Treat projects that hit the relaxed convergence band "
            "(equity +/-0.5pp, gaps <= 5x tol) as converged in the run "
            "record. Default: strict only (+/-0.25pp)."
        ),
    )
    parser.add_argument(
        "--no-save",
        dest="save_solved",
        action="store_false",
        help="Skip writing <workbook>_SOLVED.xlsm at the end of the run. Useful for fast iteration.",
    )
    parser.set_defaults(save_solved=True)
    parser.add_argument(
        "--no-output-recalc",
        dest="skip_output_recalc",
        action="store_true",
        help=(
            "Skip the final CalcOutputSheetsHL pass that recalculates Portfolio, "
            "AT Returns_WIP, Corp Model Output, Cust Prop, Dashboard, Table, and "
            "Waterfall Sensitivity. Core solve results are unaffected; Excel "
            "recalcs the output sheets lazily on the next interactive open. "
            "Saves 10-30s per run on workbooks with #REF!-heavy output sheets."
        ),
    )
    parser.set_defaults(skip_output_recalc=False)
    parser.add_argument(
        "--strip-sheets",
        default="",
        help=(
            "Comma-separated list of sheet names to DELETE from the temp copy "
            "before opening (e.g. 'Dashboard,Waterfall Sensitivity'). Useful "
            "when output sheets carry stale #REF! cells that slow SaveAs and "
            "validation. Original workbook is never modified. Only safe when "
            "no core-sheet formula references the deleted sheets."
        ),
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

    if args.show_checkpoints:
        _show_checkpoints(args.show_checkpoints)
        return

    workbook_path = Path(args.workbook) if args.workbook else DEFAULT_WORKBOOK

    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        sys.exit(1)

    strip_sheets = tuple(
        s.strip() for s in args.strip_sheets.split(",") if s.strip()
    )

    record = solve_all(
        workbook_path,
        batch_id=args.batch_id,
        dry_run=args.dry_run,
        timeout_sec=args.timeout,
        strict_validation=args.strict_validation,
        use_chunked=args.chunked,
        allow_relaxed=args.allow_relaxed,
        save_solved=args.save_solved,
        skip_output_recalc=args.skip_output_recalc,
        strip_sheets=strip_sheets,
    )

    # Exit code: 0 if converged or dry-run; 1 otherwise.
    match record.status:
        case SolveStatus.CONVERGED.value | SolveStatus.DRY_RUN.value:
            sys.exit(0)
        case _:
            sys.exit(1)


if __name__ == "__main__":
    main()
