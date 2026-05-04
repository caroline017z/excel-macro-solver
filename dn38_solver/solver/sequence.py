"""dn38_solver.solver.sequence — Per-project SolveTask construction.

Builds SolveTask objects from config constants and project info. Cell
address templates with a {col} placeholder are expanded here so the COM
runner sees fully-resolved addresses.
"""
from __future__ import annotations

from dn38_solver.types import CellAddress, ProjectInfo, SolveTask
from dn38_solver.config import (
    CELL_PROJECT_INDEX,
    READ_CELLS_TEMPLATES,
)


def _expand_address(address: str, col_letter: str) -> str:
    """Replace {col} placeholder with the project's column letter."""
    return address.replace("{col}", col_letter)


def build_solve_task(
    project: ProjectInfo,
    workbook_path: str,
) -> SolveTask:
    """Construct a fully-expanded SolveTask for one project.

    All {col} placeholders are resolved to the project's column letter.
    The resulting SolveTask is self-contained — direct_runner needs no
    knowledge of config.py.
    """
    col = project.col_letter

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
        project_index_cell=CELL_PROJECT_INDEX,
        read_cells=read_cells,
    )
