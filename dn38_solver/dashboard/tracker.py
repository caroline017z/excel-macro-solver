"""Streamlit real-time solver progress tracker.

Reads a JSON status file written by the COM runner during the solve.
Displays per-project progress bars, timing, and convergence status.

Usage:
    streamlit run dn38_solver/dashboard/tracker.py
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import streamlit as st

STATUS_FILE = Path(__file__).resolve().parent.parent.parent / "solver_status.json"

st.set_page_config(
    page_title="38DN Solver Tracker",
    page_icon="⚡",
    layout="wide",
)


def load_status() -> dict | None:
    if not STATUS_FILE.exists():
        return None
    try:
        return json.loads(STATUS_FILE.read_text())
    except (json.JSONDecodeError, OSError):
        return None


def _render_project_row(p: dict) -> None:
    """Render one project's in-progress row. Shared by flat and per-worker views."""
    name = p.get("name", "?")
    proj_status = p.get("status", "pending")
    if proj_status == "solving":
        iteration = p.get("iteration", 0)
        max_iter = p.get("max_iter", 8)
        pct = iteration / max(max_iter, 1)
        st.progress(pct, text=f"🔄 {name} — iter {iteration}/{max_iter}")
    elif proj_status == "converged":
        st.progress(1.0, text=f"✅ {name} — converged")
    elif proj_status == "not_converged":
        st.progress(1.0, text=f"⚠️ {name} — not converged")
    else:
        st.progress(0.0, text=f"⏳ {name} — pending")


def main() -> None:
    st.title("38DN Convergence Solver")

    status = load_status()

    if status is None:
        st.info("Waiting for solver to start... (no status file found)")
        st.caption(f"Looking for: `{STATUS_FILE}`")
        time.sleep(2)
        st.rerun()
        return

    phase = status.get("phase", "unknown")
    workbook = status.get("workbook", "—")
    projects = status.get("projects", [])
    total_projects = status.get("total_projects", len(projects))
    elapsed = status.get("elapsed_sec", 0)
    macro_name = status.get("macro_used", "—")
    worker_count = status.get("worker_count")  # set by StatusAggregator in parallel mode

    # Header
    is_parallel = worker_count is not None and worker_count > 1
    cols = st.columns(4 if is_parallel else 3)
    cols[0].metric("Phase", phase.replace("_", " ").title())
    cols[1].metric("Elapsed", f"{elapsed:.0f}s" if elapsed < 120 else f"{elapsed/60:.1f} min")
    cols[2].metric("Workbook", Path(workbook).stem[:40] if workbook else "—")
    if is_parallel:
        cols[3].metric("Workers", str(worker_count))

    st.divider()

    if phase == "complete":
        total_time = status.get("total_time_sec", elapsed)
        error = status.get("error")

        if error:
            st.error(f"Solver error: {error}")
        else:
            st.success(f"Solve complete in {total_time:.1f}s ({total_time/60:.1f} min)")

        # Results table
        if projects:
            st.subheader("Results")
            for p in projects:
                name = p.get("name", "?")
                converged = p.get("converged", False)
                npp = p.get("npp")
                dev_fee = p.get("dev_fee")
                fmv = p.get("fmv")
                eq_pct = p.get("equity_pct")
                irr_gap = p.get("irr_gap")

                icon = "✅" if converged else "⚠️"
                npp_s = f"${npp:.4f}" if npp is not None else "—"
                dev_s = f"${dev_fee:.4f}" if dev_fee is not None else "—"
                fmv_s = f"${fmv:.4f}" if fmv is not None else "—"
                eq_s = f"{eq_pct:.2%}" if eq_pct is not None else "—"

                st.markdown(f"{icon} **{name}** — NPP={npp_s}  DevFee={dev_s}  FMV={fmv_s}  Equity={eq_s}")

        # Timing breakdown
        st.subheader("Timing")
        open_t = status.get("open_time_sec", 0)
        macro_t = status.get("macro_time_sec", 0)
        read_t = status.get("read_time_sec", 0)

        cols = st.columns(4)
        cols[0].metric("Open", f"{open_t:.1f}s")
        cols[1].metric("Macro", f"{macro_t:.1f}s")
        cols[2].metric("Read", f"{read_t:.1f}s")
        cols[3].metric("Per Project", f"{macro_t/max(total_projects,1):.1f}s")
        return

    # --- In-progress view ---
    if phase == "opening":
        st.info("Opening workbook via COM...")
        with st.spinner("Loading..."):
            time.sleep(2)
        st.rerun()
        return

    if phase == "solving":
        current_project = status.get("current_project", 0)

        st.subheader(f"Solving {total_projects} projects")
        st.caption(f"Macro: {macro_name}")

        # Overall progress
        overall_pct = current_project / max(total_projects, 1)
        st.progress(overall_pct, text=f"Project {current_project}/{total_projects}")

        # Per-worker rendering when parallel mode is active.
        # When the StatusAggregator merges multiple worker JSONs, each
        # project entry carries a worker_id. Group them so Caroline can
        # see at a glance which worker is on which project and spot a
        # stalled instance instead of squinting at a flat list.
        if is_parallel and any("worker_id" in p for p in projects):
            grouped: dict[int, list[dict]] = {}
            for p in projects:
                grouped.setdefault(int(p.get("worker_id", 0)), []).append(p)

            worker_cols = st.columns(len(grouped))
            for col, wid in zip(worker_cols, sorted(grouped)):
                with col:
                    st.markdown(f"**Worker {wid}** — {len(grouped[wid])} project(s)")
                    for p in grouped[wid]:
                        _render_project_row(p)
        else:
            for p in projects:
                _render_project_row(p)

        time.sleep(1)
        st.rerun()
        return

    if phase == "reading":
        st.info(f"Reading results for {total_projects} projects...")
        time.sleep(1)
        st.rerun()
        return

    # Default: unknown phase, keep polling
    st.warning(f"Unknown phase: {phase}")
    time.sleep(2)
    st.rerun()


if __name__ == "__main__":
    main()
