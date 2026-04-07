"""
38DN Excel Macro Runner — Streamlit Dashboard
Visualise macro run history, NPP/FMV analytics, and batch summaries.
"""

import json
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Brand palette
# ---------------------------------------------------------------------------
NAVY = "#050D25"
GREEN = "#45A750"
TEAL = "#518484"
BLUE = "#1D6FA9"
INDIGO = "#212B48"
FAIL_RED = "#b83230"

BRAND_SEQUENCE = [BLUE, GREEN, TEAL, NAVY, INDIGO, "#6C8EBF", "#7BC17E", "#82ABAB"]

DEFAULT_DB = Path(__file__).parent / "results.db"

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="38DN Macro Runner Dashboard",
    page_icon="\u2600\ufe0f",
    layout="wide",
)

# ---------------------------------------------------------------------------
# CSS — matches VP Review App style
# ---------------------------------------------------------------------------
APP_CSS = """
<style>
:root {
    --text-primary: #050D25;
    --text-secondary: #212B48;
    --text-tertiary: #518484;
    --text-muted: #7a8291;
    --surface-base: #ffffff;
    --surface-raised: #f6f7f9;
    --surface-inset: #eef0f5;
    --border-subtle: rgba(5,13,37,0.06);
    --border-default: rgba(5,13,37,0.10);
    --border-emphasis: rgba(5,13,37,0.18);
    --brand-navy: #050D25;
    --brand-green: #45A750;
    --brand-teal: #518484;
    --brand-blue: #1D6FA9;
    --status-pass: #3a7d44;
    --status-pass-bg: rgba(69,167,80,0.08);
    --status-fail: #b83230;
    --status-fail-bg: rgba(184,50,48,0.06);
}

html, body, [class*="css"] {
    font-family: 'Century Gothic', 'Segoe UI', system-ui, sans-serif;
    color: var(--text-primary);
    font-size: 14px;
    -webkit-font-smoothing: antialiased;
}
h1,h2,h3,h4,h5,h6 {
    font-family: 'Century Gothic', 'Segoe UI', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: -0.01em;
    color: var(--text-primary) !important;
}

.block-container { padding-top: 1.2rem; max-width: 1400px; }

.hero-banner {
    background: linear-gradient(135deg, #050D25 0%, #1a2340 55%, #212B48 100%);
    border-radius: 6px;
    padding: 1.4rem 1.8rem;
    margin-bottom: 1.2rem;
    border-bottom: 2px solid rgba(69,167,80,0.4);
}
.hero-banner h1 {
    color: #ffffff !important;
    font-size: 1.3rem !important;
    margin: 0 !important;
    letter-spacing: 0.04em;
    font-weight: 700 !important;
}
.hero-banner p {
    color: rgba(255,255,255,0.5);
    font-size: 0.78rem;
    margin: 0.2rem 0 0 0;
    letter-spacing: 0.02em;
}

.kpi-row {
    display: flex;
    gap: 0.6rem;
    margin: 0.6rem 0 1rem 0;
    flex-wrap: wrap;
}
.kpi-card {
    flex: 1;
    min-width: 130px;
    background: var(--surface-base);
    border: 1px solid var(--border-default);
    border-radius: 4px;
    padding: 0.6rem 0.8rem;
}
.kpi-card .kpi-label {
    font-size: 0.62rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    color: var(--text-muted);
    margin-bottom: 0.15rem;
}
.kpi-card .kpi-value {
    font-family: 'Century Gothic', sans-serif;
    font-size: 1.2rem;
    font-weight: 700;
    color: var(--text-primary);
    line-height: 1.2;
}
.kpi-card .kpi-sub {
    font-size: 0.65rem;
    color: var(--text-muted);
    margin-top: 0.1rem;
}
.kpi-card.accent { border-left: 2.5px solid var(--brand-blue); }
.kpi-card.pass   { border-left: 2.5px solid var(--brand-green); }
.kpi-card.warn   { border-left: 2.5px solid var(--brand-teal); }
.kpi-card.fail   { border-left: 2.5px solid var(--status-fail); }

[data-testid="stSidebar"] {
    background: var(--surface-raised) !important;
    border-right: 1px solid var(--border-default);
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: var(--text-primary) !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.03em;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    border-bottom: 1px solid var(--border-emphasis);
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Century Gothic', sans-serif;
    font-weight: 600;
    font-size: 0.72rem;
    letter-spacing: 0.04em;
    padding: 0.5rem 0.9rem;
    color: var(--text-muted);
}
.stTabs [aria-selected="true"] {
    border-bottom: 2px solid var(--brand-teal) !important;
    color: var(--brand-teal) !important;
}

.stDataFrame { font-size: 0.82rem; }
.stDataFrame th { font-size: 0.7rem !important; text-transform: uppercase; letter-spacing: 0.03em; }

div[data-testid="stSidebar"] .stButton > button {
    background: var(--brand-teal) !important;
    color: #ffffff !important;
    border: 1px solid var(--brand-teal) !important;
    border-radius: 4px !important;
    font-family: 'Century Gothic', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.06em;
    width: 100%;
    transition: all 0.15s ease;
}
div[data-testid="stSidebar"] .stButton > button:hover {
    background: var(--brand-navy) !important;
    color: #ffffff !important;
}

.section-hdr {
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--text-tertiary);
    border-bottom: 1px solid var(--border-emphasis);
    padding-bottom: 0.35rem;
    margin: 1.2rem 0 0.6rem 0;
}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _connect(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


@st.cache_data(ttl=30, show_spinner=False)
def load_runs(db_path: str, limit: int = 5000) -> pd.DataFrame:
    """Load macro_runs table into a DataFrame."""
    conn = _connect(db_path)
    try:
        df = pd.read_sql_query(
            "SELECT * FROM macro_runs ORDER BY run_timestamp DESC LIMIT ?",
            conn,
            params=(limit,),
        )
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()

    if df.empty:
        return df

    df["run_timestamp"] = pd.to_datetime(df["run_timestamp"], errors="coerce")
    df["date"] = df["run_timestamp"].dt.date
    return df


def _status_color(status: str) -> str:
    if status == "success":
        return f"color: {GREEN}; font-weight: 700;"
    return f"color: {FAIL_RED}; font-weight: 700;"


def _style_status(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Apply per-cell styling to the status column."""
    def _apply(val):
        if val == "success":
            return f"background-color: rgba(69,167,80,0.10); color: {GREEN}; font-weight: 700;"
        return f"background-color: rgba(184,50,48,0.08); color: {FAIL_RED}; font-weight: 700;"

    return df.style.map(_apply, subset=["status"])


def _plotly_layout(fig, title: str = ""):
    """Apply 38DN brand layout to a Plotly figure."""
    fig.update_layout(
        title=dict(text=title, font=dict(family="Century Gothic", size=14, color=NAVY)),
        font=dict(family="Century Gothic", size=11, color=NAVY),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(font=dict(size=10)),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(5,13,37,0.06)")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(5,13,37,0.06)")
    return fig


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Macro Runner Dashboard")
    st.markdown("---")

    db_input = st.text_input(
        "Database path",
        value=str(DEFAULT_DB),
        help="Absolute path to the SQLite results database.",
    )
    db_path = db_input.strip()

    if st.button("Refresh Data"):
        st.cache_data.clear()

    st.markdown("---")
    st.caption(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
if not Path(db_path).exists():
    st.warning(f"Database not found at `{db_path}`. Run some macros first or check the path.")
    st.stop()

df = load_runs(db_path)

if df.empty:
    st.info("No macro runs recorded yet. Run a macro and come back.")
    st.stop()

# ---------------------------------------------------------------------------
# Hero banner
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="hero-banner">'
    "<h1>38DN Macro Runner Dashboard</h1>"
    "<p>Excel macro execution history, NPP analytics, and batch summaries</p>"
    "</div>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_history, tab_analysis, tab_batch = st.tabs(
    ["Run History", "Results Analysis", "Batch Summary"]
)

# ===== TAB 1 — Run History ==================================================
with tab_history:
    st.markdown('<div class="section-hdr">Filters</div>', unsafe_allow_html=True)

    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
    with col_f1:
        workbooks = sorted(df["workbook_name"].dropna().unique())
        sel_wb = st.multiselect("Workbook", workbooks, default=[], key="hist_wb")
    with col_f2:
        min_date = df["date"].min() if df["date"].notna().any() else datetime.today().date()
        date_from = st.date_input("From", value=min_date, key="hist_from")
    with col_f3:
        max_date = df["date"].max() if df["date"].notna().any() else datetime.today().date()
        date_to = st.date_input("To", value=max_date, key="hist_to")

    filtered = df.copy()
    if sel_wb:
        filtered = filtered[filtered["workbook_name"].isin(sel_wb)]
    filtered = filtered[
        (filtered["date"] >= date_from) & (filtered["date"] <= date_to)
    ]

    st.markdown('<div class="section-hdr">Run History</div>', unsafe_allow_html=True)
    st.caption(f"{len(filtered)} runs shown")

    display_cols = [
        "id", "run_timestamp", "workbook_name", "project_name", "macro_name",
        "status", "npp_per_w", "fmv_per_w", "dev_fee_per_w", "duration_sec",
        "batch_id",
    ]
    display_cols = [c for c in display_cols if c in filtered.columns]

    if not filtered.empty:
        styled = _style_status(filtered[display_cols])
        st.dataframe(
            styled,
            use_container_width=True,
            height=420,
            column_config={
                "run_timestamp": st.column_config.DatetimeColumn("Timestamp", format="YYYY-MM-DD HH:mm:ss"),
                "npp_per_w": st.column_config.NumberColumn("NPP ($/W)", format="%.4f"),
                "fmv_per_w": st.column_config.NumberColumn("FMV ($/W)", format="%.4f"),
                "dev_fee_per_w": st.column_config.NumberColumn("Dev Fee ($/W)", format="%.4f"),
                "duration_sec": st.column_config.NumberColumn("Duration (s)", format="%.1f"),
            },
        )

        # Expandable raw_outputs viewer
        st.markdown('<div class="section-hdr">Raw Outputs Inspector</div>', unsafe_allow_html=True)
        run_ids = filtered["id"].tolist()
        selected_id = st.selectbox("Select Run ID to inspect", run_ids, key="raw_sel")
        if selected_id is not None:
            row = filtered[filtered["id"] == selected_id].iloc[0]
            raw = row.get("raw_outputs")
            error = row.get("error_message")
            with st.expander(f"Run #{selected_id} — raw_outputs", expanded=True):
                if raw and str(raw) != "None":
                    try:
                        st.json(json.loads(raw) if isinstance(raw, str) else raw)
                    except (json.JSONDecodeError, TypeError):
                        st.code(str(raw))
                else:
                    st.caption("No raw outputs recorded.")
            if error and str(error) != "None":
                with st.expander(f"Run #{selected_id} — error_message", expanded=False):
                    st.error(str(error))


# ===== TAB 2 — Results Analysis =============================================
with tab_analysis:
    success_df = df[df["status"] == "success"].copy()

    if success_df.empty:
        st.info("No successful runs to analyse yet.")
    else:
        # --- NPP ($/W) by project across recent runs ---
        st.markdown('<div class="section-hdr">NPP ($/W) by Project — Latest Runs</div>', unsafe_allow_html=True)

        latest_per_project = (
            success_df.dropna(subset=["project_name", "npp_per_w"])
            .sort_values("run_timestamp")
            .drop_duplicates(subset=["project_name"], keep="last")
        )

        if not latest_per_project.empty:
            fig_bar = px.bar(
                latest_per_project.sort_values("npp_per_w", ascending=True),
                x="npp_per_w",
                y="project_name",
                orientation="h",
                color_discrete_sequence=[BLUE],
            )
            _plotly_layout(fig_bar, "NPP ($/W) by Project")
            fig_bar.update_traces(marker_line_width=0)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.caption("No NPP data available.")

        # --- NPP trend over time for a selected project ---
        st.markdown('<div class="section-hdr">NPP Trend Over Time</div>', unsafe_allow_html=True)

        projects_with_npp = sorted(
            success_df.dropna(subset=["project_name", "npp_per_w"])["project_name"].unique()
        )
        if projects_with_npp:
            sel_project = st.selectbox("Project", projects_with_npp, key="trend_proj")
            trend_df = (
                success_df[success_df["project_name"] == sel_project]
                .dropna(subset=["npp_per_w"])
                .sort_values("run_timestamp")
            )
            if not trend_df.empty:
                fig_line = px.line(
                    trend_df,
                    x="run_timestamp",
                    y="npp_per_w",
                    markers=True,
                    color_discrete_sequence=[TEAL],
                )
                _plotly_layout(fig_line, f"NPP ($/W) Trend — {sel_project}")
                st.plotly_chart(fig_line, use_container_width=True)
            else:
                st.caption("Not enough data points for a trend.")
        else:
            st.caption("No project-level NPP data available.")

        # --- FMV vs Dev Fee scatter ---
        st.markdown('<div class="section-hdr">FMV vs Dev Fee ($/W)</div>', unsafe_allow_html=True)

        scatter_df = success_df.dropna(subset=["fmv_per_w", "dev_fee_per_w"])
        if not scatter_df.empty:
            fig_scatter = px.scatter(
                scatter_df,
                x="dev_fee_per_w",
                y="fmv_per_w",
                color="project_name",
                hover_data=["workbook_name", "run_timestamp", "npp_per_w"],
                color_discrete_sequence=BRAND_SEQUENCE,
            )
            _plotly_layout(fig_scatter, "FMV ($/W) vs Dev Fee ($/W)")
            fig_scatter.update_traces(marker=dict(size=8, line=dict(width=0.5, color=NAVY)))
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.caption("No FMV / Dev Fee data available.")


# ===== TAB 3 — Batch Summary ================================================
with tab_batch:
    batches = df.dropna(subset=["batch_id"])

    if batches.empty:
        st.info("No batch runs recorded yet. Use batch mode to see aggregates here.")
    else:
        batch_ids = sorted(batches["batch_id"].unique(), reverse=True)

        st.markdown('<div class="section-hdr">Batch Overview</div>', unsafe_allow_html=True)

        # Build aggregate stats per batch
        agg = (
            batches.groupby("batch_id")
            .agg(
                total=("id", "count"),
                successes=("status", lambda s: (s == "success").sum()),
                avg_npp=("npp_per_w", "mean"),
                min_ts=("run_timestamp", "min"),
                max_ts=("run_timestamp", "max"),
            )
            .reset_index()
        )
        agg["success_rate"] = (agg["successes"] / agg["total"] * 100).round(1)
        agg["avg_npp"] = agg["avg_npp"].round(4)
        agg = agg.sort_values("max_ts", ascending=False)

        # KPI row for most recent batch
        latest_batch = agg.iloc[0]
        kpi_html = f"""
        <div class="kpi-row">
            <div class="kpi-card accent">
                <div class="kpi-label">Latest Batch</div>
                <div class="kpi-value" style="font-size:0.9rem;">{latest_batch['batch_id']}</div>
            </div>
            <div class="kpi-card pass">
                <div class="kpi-label">Workbooks Processed</div>
                <div class="kpi-value">{int(latest_batch['total'])}</div>
            </div>
            <div class="kpi-card {'pass' if latest_batch['success_rate'] >= 90 else 'fail'}">
                <div class="kpi-label">Success Rate</div>
                <div class="kpi-value">{latest_batch['success_rate']}%</div>
            </div>
            <div class="kpi-card warn">
                <div class="kpi-label">Avg NPP ($/W)</div>
                <div class="kpi-value">{latest_batch['avg_npp'] if pd.notna(latest_batch['avg_npp']) else 'N/A'}</div>
            </div>
        </div>
        """
        st.markdown(kpi_html, unsafe_allow_html=True)

        # Batch table
        st.dataframe(
            agg.rename(columns={
                "batch_id": "Batch ID",
                "total": "Total Runs",
                "successes": "Successes",
                "success_rate": "Success %",
                "avg_npp": "Avg NPP ($/W)",
                "min_ts": "Started",
                "max_ts": "Ended",
            }),
            use_container_width=True,
            hide_index=True,
        )

        # Detail drill-down
        st.markdown('<div class="section-hdr">Batch Detail</div>', unsafe_allow_html=True)
        sel_batch = st.selectbox("Select Batch", batch_ids, key="batch_sel")

        batch_detail = batches[batches["batch_id"] == sel_batch].sort_values("run_timestamp")
        detail_cols = [
            "id", "workbook_name", "project_name", "status",
            "npp_per_w", "fmv_per_w", "dev_fee_per_w", "duration_sec",
            "error_message",
        ]
        detail_cols = [c for c in detail_cols if c in batch_detail.columns]

        st.dataframe(
            _style_status(batch_detail[detail_cols]),
            use_container_width=True,
            hide_index=True,
            column_config={
                "npp_per_w": st.column_config.NumberColumn("NPP ($/W)", format="%.4f"),
                "fmv_per_w": st.column_config.NumberColumn("FMV ($/W)", format="%.4f"),
                "dev_fee_per_w": st.column_config.NumberColumn("Dev Fee ($/W)", format="%.4f"),
                "duration_sec": st.column_config.NumberColumn("Duration (s)", format="%.1f"),
            },
        )
