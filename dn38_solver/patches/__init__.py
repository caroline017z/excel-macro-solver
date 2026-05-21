"""dn38_solver.patches — Idempotent input-workbook patches.

Patches run BEFORE the solve to repair known template-cruft patterns that
silently produce formula errors in summary tabs without affecting the
solver convergence path. Each patch is safe to re-run; calling it on a
clean workbook is a no-op.
"""
