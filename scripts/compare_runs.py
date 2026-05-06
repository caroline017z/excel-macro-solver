"""scripts.compare_runs — Side-by-side baseline vs. variant solver run.

DECISION RULE: Flip ``USE_TIGHT_NPP_SCOPE = True`` in SolveHeadless.bas iff the
variant run shows a portfolio-total-time speedup of >= 10% (Speedup >= 1.10x)
AND the variant introduces zero new non-convergences relative to baseline.

Usage:
    python -m scripts.compare_runs <baseline_solved_xlsx> <variant_solved_xlsx>
    python scripts/compare_runs.py <baseline_solved_xlsx> <variant_solved_xlsx>

Both inputs are ``*_SOLVED.xlsx`` files produced by the headless macro for the
same portfolio. Projects are matched by name (column B). Rows present on only
one side are surfaced in a MISMATCH section so a partial run is obvious.

Output columns:
    Baseline / Variant : per-project total Solve Seconds (column M)
    Speedup            : baseline / variant (>1 means variant is faster)
    DSCR Δ% / NPP Δ%   : variant phase seconds relative to baseline phase
    Appr Δ%              seconds; (variant - baseline) / baseline.

The script does not auto-recommend the toggle flip — read the portfolio
Speedup line and the convergence-rate comparison against the rule above.
"""
from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass
from pathlib import Path

# Support both `python -m scripts.compare_runs ...` (package import works
# directly) and `python scripts/compare_runs.py ...` (no package context;
# fall back to a sibling-file import by adding the script's directory to
# sys.path).
try:
    from scripts.analyze_solver_results import ProjectRow, load_results
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from analyze_solver_results import ProjectRow, load_results  # type: ignore[no-redef]

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class PairedResult:
    """Baseline + variant rows for a single project, matched by name."""

    name: str
    baseline: ProjectRow
    variant: ProjectRow

    @property
    def delta_secs(self) -> float:
        return self.variant.solve_secs - self.baseline.solve_secs

    @property
    def speedup(self) -> float | None:
        if self.variant.solve_secs <= 0:
            return None
        return self.baseline.solve_secs / self.variant.solve_secs


def _index_by_name(rows: list[ProjectRow]) -> dict[str, ProjectRow]:
    """Index rows by project name. Duplicate names are logged and last-write-wins."""
    out: dict[str, ProjectRow] = {}
    for r in rows:
        if r.name in out:
            log.warning("Duplicate project name in run: %s — keeping last row", r.name)
        out[r.name] = r
    return out


def pair_runs(
    baseline: list[ProjectRow],
    variant: list[ProjectRow],
) -> tuple[list[PairedResult], list[ProjectRow], list[ProjectRow]]:
    """Match baseline and variant rows by project name.

    Returns (paired, baseline_only, variant_only).
    """
    b_idx = _index_by_name(baseline)
    v_idx = _index_by_name(variant)

    paired: list[PairedResult] = []
    for name in b_idx:
        if name in v_idx:
            paired.append(PairedResult(name=name, baseline=b_idx[name], variant=v_idx[name]))

    # Preserve original input order for the "only" lists so the report
    # mirrors what the user saw on disk.
    baseline_only = [r for r in baseline if r.name not in v_idx]
    variant_only = [r for r in variant if r.name not in b_idx]
    return paired, baseline_only, variant_only


def _delta_pct(baseline_v: float, variant_v: float) -> str:
    """Format (variant - baseline) / baseline as a signed 1-decimal percent."""
    if baseline_v <= 0:
        if variant_v <= 0:
            return "  -  "
        return "  +inf"
    pct = (variant_v - baseline_v) / baseline_v * 100
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:5.1f}%"


def _speedup_str(s: float | None) -> str:
    if s is None:
        return "   -   "
    return f"{s:5.2f}x"


def render_table(
    paired: list[PairedResult],
    baseline_only: list[ProjectRow],
    variant_only: list[ProjectRow],
) -> str:
    header = (
        f"{'Project':<32} "
        f"{'Baseline':>9} {'Variant':>9} {'Δ (s)':>8} "
        f"{'Speedup':>8} "
        f"{'DSCR Δ%':>8} {'NPP Δ%':>8} {'Appr Δ%':>8}"
    )
    sep = (
        f"{'-'*32} "
        f"{'-'*9} {'-'*9} {'-'*8} "
        f"{'-'*8} "
        f"{'-'*8} {'-'*8} {'-'*8}"
    )
    out: list[str] = [header, sep]

    for p in paired:
        out.append(
            f"{p.name[:32]:<32} "
            f"{p.baseline.solve_secs:9.1f} "
            f"{p.variant.solve_secs:9.1f} "
            f"{p.delta_secs:+8.1f} "
            f"{_speedup_str(p.speedup):>8} "
            f"{_delta_pct(p.baseline.dscr_secs, p.variant.dscr_secs):>8} "
            f"{_delta_pct(p.baseline.npp_secs, p.variant.npp_secs):>8} "
            f"{_delta_pct(p.baseline.appr_secs, p.variant.appr_secs):>8}"
        )

    # Portfolio aggregate.
    b_total = sum(p.baseline.solve_secs for p in paired)
    v_total = sum(p.variant.solve_secs for p in paired)
    delta_total = v_total - b_total
    speedup = (b_total / v_total) if v_total > 0 else None

    b_dscr = sum(p.baseline.dscr_secs for p in paired)
    v_dscr = sum(p.variant.dscr_secs for p in paired)
    b_npp = sum(p.baseline.npp_secs for p in paired)
    v_npp = sum(p.variant.npp_secs for p in paired)
    b_appr = sum(p.baseline.appr_secs for p in paired)
    v_appr = sum(p.variant.appr_secs for p in paired)

    out.append(sep)
    out.append(
        f"{'PORTFOLIO':<32} "
        f"{b_total:9.1f} "
        f"{v_total:9.1f} "
        f"{delta_total:+8.1f} "
        f"{_speedup_str(speedup):>8} "
        f"{_delta_pct(b_dscr, v_dscr):>8} "
        f"{_delta_pct(b_npp, v_npp):>8} "
        f"{_delta_pct(b_appr, v_appr):>8}"
    )

    if baseline_only or variant_only:
        out.append("")
        out.append("=" * 88)
        out.append("  MISMATCH — projects present in only one run")
        out.append("=" * 88)
        if baseline_only:
            out.append("  Baseline only:")
            for r in baseline_only:
                out.append(f"    - {r.name}  ({r.solve_secs:.1f}s)")
        if variant_only:
            out.append("  Variant only:")
            for r in variant_only:
                out.append(f"    - {r.name}  ({r.solve_secs:.1f}s)")

    return "\n".join(out)


def render_summary(
    paired: list[PairedResult],
    baseline_only: list[ProjectRow],
    variant_only: list[ProjectRow],
) -> str:
    n = len(paired)
    if n == 0:
        return "\nNo overlapping projects between the two runs.\n"

    b_total = sum(p.baseline.solve_secs for p in paired)
    v_total = sum(p.variant.solve_secs for p in paired)
    delta_total = v_total - b_total
    speedup = (b_total / v_total) if v_total > 0 else None
    speedup_pct = (speedup - 1) * 100 if speedup is not None else None

    b_conv = sum(1 for p in paired if p.baseline.converged)
    v_conv = sum(1 for p in paired if p.variant.converged)
    new_failures = [
        p.name for p in paired if p.baseline.converged and not p.variant.converged
    ]
    recovered = [
        p.name for p in paired if not p.baseline.converged and p.variant.converged
    ]

    lines: list[str] = []
    lines.append("")
    lines.append("=" * 88)
    lines.append("  Comparison Summary")
    lines.append("=" * 88)
    lines.append(f"  Paired projects         : {n}")
    lines.append(f"  Baseline-only projects  : {len(baseline_only)}")
    lines.append(f"  Variant-only projects   : {len(variant_only)}")
    lines.append("")
    lines.append(f"  Total Solve Seconds")
    lines.append(f"    Baseline              : {b_total:.1f}")
    lines.append(f"    Variant               : {v_total:.1f}")
    lines.append(f"    Δ                     : {delta_total:+.1f}")
    if speedup is not None and speedup_pct is not None:
        lines.append(f"    Speedup               : {speedup:.2f}x  ({speedup_pct:+.1f}%)")
    else:
        lines.append("    Speedup               : -")
    lines.append("")
    lines.append(f"  Convergence (paired)")
    lines.append(f"    Baseline              : {b_conv}/{n} ({(b_conv/n)*100:.1f}%)")
    lines.append(f"    Variant               : {v_conv}/{n} ({(v_conv/n)*100:.1f}%)")
    lines.append(f"    New non-convergences  : {len(new_failures)}"
                 + (f"  -> {', '.join(new_failures)}" if new_failures else ""))
    lines.append(f"    Newly converging      : {len(recovered)}"
                 + (f"  -> {', '.join(recovered)}" if recovered else ""))
    lines.append("")
    lines.append("  Decision rule (see module docstring):")
    lines.append("    Flip USE_TIGHT_NPP_SCOPE = True iff Speedup >= 1.10x AND")
    lines.append("    New non-convergences == 0.")
    lines.append("")

    return "\n".join(lines)


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="compare_runs",
        description=(
            "Compare two _SOLVED.xlsx runs side by side to evaluate "
            "phase-scope toggle changes."
        ),
    )
    parser.add_argument("baseline", type=Path, help="Baseline _SOLVED.xlsx path.")
    parser.add_argument("variant", type=Path, help="Variant _SOLVED.xlsx path.")
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable DEBUG logging.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
        stream=sys.stderr,
    )

    baseline_rows = load_results(args.baseline)
    variant_rows = load_results(args.variant)

    if not baseline_rows:
        print(f"No project rows in baseline: {args.baseline}")
        return 1
    if not variant_rows:
        print(f"No project rows in variant: {args.variant}")
        return 1

    paired, b_only, v_only = pair_runs(baseline_rows, variant_rows)

    print(f"Baseline : {args.baseline}")
    print(f"Variant  : {args.variant}")
    print()
    print(render_table(paired, b_only, v_only))
    print(render_summary(paired, b_only, v_only))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
