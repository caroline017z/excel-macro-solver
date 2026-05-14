"""dn38_solver.com.hresult — Decode COM HRESULTs into actionable errors.

Excel raises generic `com_error: (-2147352567, 'Exception occurred.', ...)`
for a wide range of internal failures. Without decoding, the orchestrator
surfaces opaque tuples that take minutes to diagnose.

This module pattern-matches common COM exception shapes and returns a
single-line human-readable hint + a flag indicating whether automatic
recovery is worth attempting.

Verified failure modes (2026-05-13 RP Puma incident):
- DISP_E_EXCEPTION (-2147352567 / 0x80020009) with secondary -2146788248
  (0x800a9c68): generic VBA runtime error inside Application.Run.
  Almost always means the workbook state is stale relative to the macro.
  Recovery: re-import the .bas via Excel COM and retry once.
"""
from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class DecodedComError:
    """Human-readable decode of a COM exception."""
    raw: str
    hresult: int | None
    secondary: int | None        # The nested HRESULT in the excepinfo tuple
    summary: str                 # One-line summary for log output
    likely_cause: str            # Best-guess root cause
    recovery: str                # Suggested operator action (one line)
    auto_recoverable: bool       # If True, orchestrator may retry once after re-import


# Pattern matches `(-2147352567, 'Exception occurred.', (..., <secondary>), None)`
_COM_ERR_RE = re.compile(
    r"\((-?\d+),\s*'([^']*)',\s*\(([^)]*)\),\s*([^)]*)\)"
)


def _parse_excepinfo(s: str) -> int | None:
    """Pull the secondary HRESULT (6th element of the excepinfo tuple) out
    of the parenthesized inner group. The standard pywin32 excepinfo is
    `(code, source, description, helpfile, helpcontext, scode)` — scode
    is the COM error that triggered the wrapper.
    """
    parts = [p.strip() for p in s.split(",")]
    if len(parts) >= 6:
        try:
            return int(parts[5])
        except ValueError:
            return None
    return None


def decode_com_error(raw_exception: str) -> DecodedComError:
    """Pattern-match a stringified pywin32 com_error tuple and return
    a structured decode. Always returns a result — falls back to
    "unknown COM error" if no pattern matches.
    """
    m = _COM_ERR_RE.search(raw_exception)
    if not m:
        return DecodedComError(
            raw=raw_exception,
            hresult=None,
            secondary=None,
            summary="COM exception (could not parse HRESULT)",
            likely_cause="Unrecognized error shape; check raw output",
            recovery="Inspect Excel manually; check log for surrounding context",
            auto_recoverable=False,
        )

    hresult = int(m.group(1))
    excepinfo_inner = m.group(3)
    secondary = _parse_excepinfo(excepinfo_inner)

    # -2147352567 = 0x80020009 = DISP_E_EXCEPTION (the wrapper for "an
    # exception occurred inside the COM method call"). Always look at
    # the secondary code for the real failure.
    if hresult == -2147352567:
        if secondary == -2146788248:  # 0x800a9c68 — generic Office automation error
            return DecodedComError(
                raw=raw_exception,
                hresult=hresult,
                secondary=secondary,
                summary="Excel reports 'Exception occurred' inside the VBA macro",
                likely_cause=(
                    "Generic VBA runtime error during macro execution. "
                    "Common when the workbook's state has been corrupted by a "
                    "non-Excel writer (e.g., openpyxl save on .xlsm) since the "
                    "last macro import."
                ),
                recovery=(
                    "Re-import the macro via Excel COM "
                    "(python import_vba_module.py <workbook>) and retry. "
                    "If that doesn't help, open the workbook in Excel + "
                    "Alt+F11 to see the real VBA error line."
                ),
                auto_recoverable=True,
            )
        if secondary == -2147418113:  # 0x8000FFFF E_UNEXPECTED
            return DecodedComError(
                raw=raw_exception,
                hresult=hresult,
                secondary=secondary,
                summary="Excel reports E_UNEXPECTED inside the macro",
                likely_cause=(
                    "Excel object model failure mid-call. Often Excel process "
                    "instability, file lock contention, or a corrupted workbook."
                ),
                recovery=(
                    "Kill all Excel processes; reopen the workbook fresh in a "
                    "new orchestrator run."
                ),
                auto_recoverable=False,  # E_UNEXPECTED retries rarely help
            )
        return DecodedComError(
            raw=raw_exception,
            hresult=hresult,
            secondary=secondary,
            summary=f"Excel macro raised an exception (secondary HRESULT {secondary})",
            likely_cause="VBA runtime error; specific code not in decoder table",
            recovery=(
                "Re-import the macro and retry. If repeated, open in Excel and "
                "debug via Alt+F11."
            ),
            auto_recoverable=True,
        )

    if hresult == -2147023174:  # 0x800706BA "RPC server unavailable"
        return DecodedComError(
            raw=raw_exception,
            hresult=hresult,
            secondary=secondary,
            summary="Excel COM RPC server unavailable (Excel process died)",
            likely_cause=(
                "Excel.exe exited mid-call. Most often an in-process crash "
                "or the ~900s RPC timeout firing on a long uninterrupted "
                "macro call."
            ),
            recovery=(
                "Retry with --chunked (per-project macro calls stay under the "
                "RPC timeout) or --workers N (parallel processes; each is its "
                "own Excel instance)."
            ),
            auto_recoverable=False,
        )

    if hresult == -2147417848:  # 0x80010108 "object disconnected from clients"
        return DecodedComError(
            raw=raw_exception,
            hresult=hresult,
            secondary=secondary,
            summary="Excel object disconnected (process exited)",
            likely_cause="Excel.exe shut down mid-orchestration. Often heap exhaustion.",
            recovery="Close all Excel; rerun with fewer workers or --no-output-recalc.",
            auto_recoverable=False,
        )

    return DecodedComError(
        raw=raw_exception,
        hresult=hresult,
        secondary=secondary,
        summary=f"COM error with HRESULT {hresult} (no decoder entry)",
        likely_cause="Unknown — see raw exception",
        recovery="Inspect Excel log; surface the raw error to the maintainer",
        auto_recoverable=False,
    )


def format_decoded(d: DecodedComError) -> str:
    """Render a DecodedComError as a multi-line block for log output."""
    lines = [
        f"  COM error: {d.summary}",
        f"    HRESULT:    {hex(d.hresult & 0xFFFFFFFF) if d.hresult is not None else 'n/a'}"
        f"  ({d.hresult})",
    ]
    if d.secondary is not None:
        lines.append(
            f"    Secondary:  {hex(d.secondary & 0xFFFFFFFF)}  ({d.secondary})"
        )
    lines.append(f"    Likely:     {d.likely_cause}")
    lines.append(f"    Recovery:   {d.recovery}")
    if d.auto_recoverable:
        lines.append(
            f"    Auto-recovery: ELIGIBLE — orchestrator will attempt re-import + retry once"
        )
    return "\n".join(lines)
