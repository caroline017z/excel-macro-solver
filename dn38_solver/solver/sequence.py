"""dn38_solver.solver.sequence — GoalSeek task construction.

Builds SolveTask objects from config constants and project info.
All cell address templates are expanded here, not in the COM worker.
"""
from __future__ import annotations

from dn38_solver.types import CellAddress, GoalSeekOp, ProjectInfo, SolveTask
from dn38_solver.config import (
    CALC_CORE_SHEETS,
    CELL_EQUITY,
    CELL_EQUITY_TARGET,
    CELL_HOLDCO_TOGGLE,
    CELL_PROJECT_INDEX,
    CELL_TOTAL_USES,
    EQUITY_TOLERANCE,
    GOALSEEK_PHASE1,
    GOALSEEK_PHASE2_TEMPLATES,
    GS_MAX_CHANGE,
    GS_MAX_ITERATIONS,
    IRR_TOLERANCE,
    MAX_GS_RETRIES,
    MAX_ITERATIONS,
    READ_CELLS_TEMPLATES,
)


def _expand_address(address: str, col_letter: str) -> str:
    """Replace {col} placeholder with the project's column letter."""
    return address.replace("{col}", col_letter)


def _expand_goalseek(op: GoalSeekOp, col_letter: str) -> GoalSeekOp:
    """Expand {col} placeholders in a GoalSeekOp template."""
    return GoalSeekOp(
        target_sheet=op.target_sheet,
        target_cell=_expand_address(op.target_cell, col_letter),
        goal_sheet=op.goal_sheet,
        goal_cell=_expand_address(op.goal_cell, col_letter),
        changing_sheet=op.changing_sheet,
        changing_cell=_expand_address(op.changing_cell, col_letter),
        lower_bound=op.lower_bound,
        upper_bound=op.upper_bound,
    )


def build_solve_task(
    project: ProjectInfo,
    workbook_path: str,
) -> SolveTask:
    """Construct a fully-expanded SolveTask for one project.

    All {col} placeholders are resolved to the project's column letter.
    The resulting SolveTask is self-contained — the COM worker needs
    no knowledge of config.py.
    """
    col = project.col_letter

    phase2 = tuple(
        _expand_goalseek(op, col) for op in GOALSEEK_PHASE2_TEMPLATES
    )

    read_cells = tuple(
        CellAddress(
            sheet=cell.sheet,
            address=_expand_address(cell.address, col),
        )
        for cell in READ_CELLS_TEMPLATES
    )

    return SolveTask(
        workbook_path=workbook_path,
        project_offset=project.offset,
        project_col_letter=col,
        project_name=project.name,
        calc_core_sheets=CALC_CORE_SHEETS,
        project_index_cell=CELL_PROJECT_INDEX,
        holdco_cell=CELL_HOLDCO_TOGGLE,
        equity_cell=CELL_EQUITY,
        equity_target_cell=CELL_EQUITY_TARGET,
        total_uses_cell=CELL_TOTAL_USES,
        goal_seeks_phase1=GOALSEEK_PHASE1,
        goal_seeks_phase2=phase2,
        max_iterations=MAX_ITERATIONS,
        max_gs_retries=MAX_GS_RETRIES,
        irr_tolerance=IRR_TOLERANCE,
        equity_tolerance=EQUITY_TOLERANCE,
        gs_max_change=GS_MAX_CHANGE,
        gs_max_iterations=GS_MAX_ITERATIONS,
        read_cells=read_cells,
    )
