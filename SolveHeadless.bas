Attribute VB_Name = "modSolveHeadless"
' ==============================================================================
'  HEADLESS CONVERGENCE SOLVER v5 — for Python COM automation
'
'  Performance optimizations:
'    1. No MsgBox / StatusBar / user dialogs
'    2. Leaves Calculation in xlCalculationManual on exit
'    3. Deterministic recalc ladder:
'         - Tier 1: Application.Calculate
'         - Tier 2: Application.Calculate + Application.Calculate
'         - Tier 3: Application.CalculateFull
'    4. Non-core sheets disabled via EnableCalculation = False
'    5. Sequential GoalSeeks with recalc after each operation
'    6. Multi-threaded calculation enabled for parallel formula evaluation
'    7. Adaptive warm/cold GoalSeek settings per project
' ==============================================================================

Option Explicit

' --- Configuration ---
' Outer-loop budget: most projects converge in 1-3 iters or never. Iterating
' past MAX_ITER_SOFT with no equity progress is wasted recalcs, so we exit
' early once equity has settled. MAX_ITER_HARD is the cap for projects that
' are still moving meaningfully and just need a few more passes.
Private Const MAX_ITER_SOFT     As Integer = 4
Private Const MAX_ITER_HARD     As Integer = 8
Private Const SOFT_CAP_PROGRESS_TOL As Double = 0.001   ' 0.1pp equity move
Private Const MAX_ITER          As Integer = 8          ' = MAX_ITER_HARD; loop bound
Private Const MAX_GS_RETRY_WARM As Integer = 3
Private Const MAX_GS_RETRY_COLD As Integer = 6
Private Const GS_MAXITER_WARM   As Integer = 200
Private Const GS_MAXITER_COLD   As Integer = 1000
Private Const EQUITY_FINAL_TOL  As Double = 0.0025   ' +/-0.25pp = 25bps off 10% min equity target
Private Const EQUITY_RELAXED_TOL As Double = 0.005   ' +/-0.5pp band; investment-grade but outside strict
Private Const RELAXED_GAP_FACTOR As Double = 5#       ' inner gaps must be <= 5x tol to count as relaxed
Private Const IRR_TOLERANCE     As Double = 0.0003
Private Const APPR_TOLERANCE    As Double = 0.0003
Private Const DSCR_MIN          As Double = 0.25
Private Const DSCR_MAX          As Double = 3#
Private Const PROJECT_TIMEOUT_SECONDS As Double = 1200   ' 20 min hard cap per project — catches runaway tier-3 escalations
Private Const COL_SCAN_LIMIT    As Integer = 60

' Sanity bounds for NPP / Dev Fee — GoalSeek is unconstrained Newton-style
' and can diverge to absurd values (e.g. $12/W) when local slope is small.
' Out-of-range solves are snapped to the seed for the next retry.
Private Const NPP_MIN           As Double = -0.2
Private Const NPP_MAX           As Double = 2.0
Private Const NPP_SEED          As Double = 0.2
Private Const DEV_FEE_MIN       As Double = 0.05
Private Const DEV_FEE_MAX       As Double = 5.5
Private Const DEV_FEE_SEED      As Double = 0.2

' --- Sheet names ---
Private Const SHT_PI  As String = "Project Inputs"
Private Const SHT_PT  As String = "PT Returns"
Private Const SHT_SOLVER_RESULTS As String = "__SolverResults"

' --- Cell addresses ---
Private Const PT_HOLDCO_ONOFF   As String = "C134"
Private Const PT_EQUITY         As String = "C128"
Private Const PT_MIN_EQ_TARGET  As String = "F128"
Private Const PT_DSCR_MULTIPLE  As String = "F129"
Private Const PT_TOTAL_USES     As String = "C130"
Private Const PI_PROJ_INDEX     As String = "F2"
Private Const PI_IRR_LIVE       As String = "F37"
Private Const PI_IRR_TARGET     As String = "F36"
Private Const PI_APPR_LIVE      As String = "F31"
Private Const PI_WACC_TARGET    As String = "F30"

' --- Row/column layout ---
Private Const PI_ROW_TOGGLE     As Long = 7
Private Const PI_ROW_NAME       As Long = 4
Private Const PI_ROW_NPP        As Long = 38
Private Const PI_ROW_DEV_FEE    As Long = 32
Private Const PI_FIRST_PROJ_COL As Integer = 8
Private Const PI_BASE_COL       As Integer = 7

' --- Recalc ladder state ---
Private mCalcTier As Integer

' --- Phase-specific recalc scopes ---
' Each major step of the solve loop touches only a subset of the model.
' Tier 1 dispatches to a phase-specific scope; tiers 2/3 fall back to wider
' coverage so the deterministic recalc ladder still terminates correctly.
'
' Set USE_TIGHT_*_SCOPE = False to revert that phase to the full 13-sheet
' recalc and rule it out as a cause if convergence regresses.
Private Const PHASE_FULL As Integer = 0
Private Const PHASE_DSCR As Integer = 1
Private Const PHASE_NPP  As Integer = 2
Private Const PHASE_APPR As Integer = 3

Private Const USE_TIGHT_DSCR_SCOPE As Boolean = True
Private Const USE_TIGHT_NPP_SCOPE  As Boolean = False
Private Const USE_TIGHT_APPR_SCOPE As Boolean = True

' --- Per-project phase timing telemetry (seconds, accumulated per project) ---
Private mCalcSecsDSCR As Double
Private mCalcSecsNPP  As Double
Private mCalcSecsAppr As Double
Private mCalcSecsFull As Double

' --- Output-sheet recalc toggle (Python-controllable) ---
' When True, FinalizeSolveEnvHL / SolveHeadless skip CalcOutputSheetsHL,
' leaving downstream sheets (Dashboard, Waterfall Sensitivity, etc.)
' un-recalculated. Excel recalcs them lazily on next interactive open.
' Default False preserves prior behavior.
Private mSkipOutputRecalc As Boolean


Private Function EnsureSolverResultsSheetHL() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_SOLVER_RESULTS)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHT_SOLVER_RESULTS
    End If

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0

    Set EnsureSolverResultsSheetHL = ws
End Function

Private Sub ResetSolverResultsHL(ByVal wsRes As Worksheet)
    wsRes.Cells.ClearContents
    ' ClearContents preserves NumberFormat from prior runs. Pin cols O-R
    ' (per-phase calc seconds) to 4-decimal so the values match what
    ' Round(..., 4) writes -- a stale "%" or "$" format on these cells
    ' would silently misread the telemetry in __SolverResults.
    wsRes.Columns("O:R").NumberFormat = "0.0000"
    wsRes.Range("A1").Value = "Project Offset"
    wsRes.Range("B1").Value = "Project Name"
    wsRes.Range("C1").Value = "DSCR"
    wsRes.Range("D1").Value = "NPP"
    wsRes.Range("E1").Value = "Dev Fee"
    wsRes.Range("F1").Value = "Equity Pct"
    wsRes.Range("G1").Value = "IRR Gap"
    wsRes.Range("H1").Value = "Appraisal Gap"
    wsRes.Range("I1").Value = "Converged"
    wsRes.Range("J1").Value = "Calc Tier"
    wsRes.Range("K1").Value = "GS Retry Limit"
    wsRes.Range("L1").Value = "Mode"
    wsRes.Range("M1").Value = "Solve Seconds"
    wsRes.Range("N1").Value = "Heartbeat UTC"
    wsRes.Range("O1").Value = "Calc Secs DSCR"
    wsRes.Range("P1").Value = "Calc Secs NPP"
    wsRes.Range("Q1").Value = "Calc Secs Appr"
    wsRes.Range("R1").Value = "Calc Secs Full"
    wsRes.Range("S1").Value = "Iterations"
    wsRes.Range("T1").Value = "Conv Tier"
End Sub

' Classify a project's outcome based on its final equity, IRR gap, and Appr gap.
' Strict: equity within +/-0.25pp AND both inner gaps within tight tolerance.
' Relaxed: equity within +/-0.5pp AND both inner gaps within 5x tolerance.
' None: anything else. Strict beats relaxed; relaxed beats none.
'
' This is a classification helper only -- it does not affect bConverged.
' Column I (Converged) keeps its strict-only semantics; column T carries
' the tier so Python can apply --allow-relaxed policy at the run level.
Private Function ClassifyConvergenceHL(ByVal dEquityPct As Double, _
                                       ByVal dIRRGap As Double, _
                                       ByVal dApprGap As Double) As String
    If Abs(dEquityPct - 0.1) <= EQUITY_FINAL_TOL _
       And dIRRGap <= IRR_TOLERANCE _
       And dApprGap <= APPR_TOLERANCE Then
        ClassifyConvergenceHL = "strict"
    ElseIf Abs(dEquityPct - 0.1) <= EQUITY_RELAXED_TOL _
       And dIRRGap <= IRR_TOLERANCE * RELAXED_GAP_FACTOR _
       And dApprGap <= APPR_TOLERANCE * RELAXED_GAP_FACTOR Then
        ClassifyConvergenceHL = "relaxed"
    Else
        ClassifyConvergenceHL = "none"
    End If
End Function

Private Sub WriteHeartbeatHL(ByVal wsRes As Worksheet, ByVal heartbeatText As String)
    wsRes.Range("N1").Value = heartbeatText
End Sub

' ==============================================================================
'  INTERNAL: Safe formula -> constant via cell-self-assign.
'  Refuses to stamp Excel error values (#DIV/0!, #VALUE!, #N/A, etc.) into
'  the target cell. Without this guard the self-assign would freeze the
'  error string into the saved workbook; the post-merge sanity gate then
'  string-skips it (non-numeric expected -> silent skip in earlier code,
'  now a flagged finding) and Caroline ships an IC-facing column with
'  #DIV/0! visible. On error encounter we leave the formula in place and
'  append an audit row to __SolverResults so the operator can trace which
'  cell was skipped and which project / call site was responsible.
' ==============================================================================
Private Sub HardStampNumericHL(ByVal ws As Worksheet, _
                                ByVal r As Long, _
                                ByVal c As Long, _
                                ByVal context As String)
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsError(v) Then
        ' Audit-row write must NEVER abort the solve. If __SolverResults
        ' is structure-protected, the workbook is read-only, or
        ' Worksheets.Add fails (workbook protection, sheet limit, etc.),
        ' an unhandled error here would propagate up through
        ' SolveOneProjectByColHL / SolveHeadless and kill the run mid-
        ' portfolio — the exact failure mode the helper is meant to
        ' prevent. Bracket the entire audit-row block so any failure to
        ' log just drops the audit silently; the formula-not-stamped
        ' invariant (which is what protects the IC-facing column from
        ' #DIV/0!) still holds because we Exit Sub before the value
        ' assignment regardless.
        On Error Resume Next
        Dim wsRes As Worksheet
        Set wsRes = EnsureSolverResultsSheetHL()
        If Not wsRes Is Nothing Then
            Dim auditRow As Long
            auditRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row + 1
            wsRes.Cells(auditRow, 2).Value = "stamp_skipped:" & context & _
                ":r" & r & "c" & c & ":" & CStr(v)
            wsRes.Cells(auditRow, 14).Value = CStr(Now)
        End If
        On Error GoTo 0
        Exit Sub
    End If
    ws.Cells(r, c).Value = v
End Sub


' ==============================================================================
'  INTERNAL HELPERS
' ==============================================================================
Private Sub CalcModelCoreHL()
    ' Backwards-compatible alias for the original full-scope recalc.
    ' Use CalcForPhase(PHASE_*) for phase-aware call sites; this entry
    ' point preserves identical behaviour for setup / restore / fallback.
    CalcForPhase PHASE_FULL
End Sub

Private Sub CalcForPhase(ByVal phase As Integer)
    ' Deterministic recalc ladder, dispatched by solve-loop phase:
    '   Tier 1 = phase-specific subset of sheets (fastest correct path)
    '   Tier 2 = tier 1 twice, for deeper propagation
    '   Tier 3 = CalculateFull guardrail for cold-start edge cases
    '
    ' Per-sheet Sheets("X").Calculate is required because Application.Calculate
    ' under xlCalculationManual + multi-threaded calc does not reliably mark
    ' cross-sheet OFFSET-via-F2 dependencies dirty.
    '
    ' Phase scopes are chosen conservatively. DSCR change ripples through the
    ' capital stack (Tax Equity, Perm Debt, CL) into PT Returns; NPP change
    ' affects almost everything (full scope by default); Appraisal change is
    ' local to the Appraisal sheet's Dev-Fee-driven valuation.
    Dim t0 As Double
    t0 = Timer

    Select Case mCalcTier
        Case 1
            CalcPhaseScope phase
        Case 2
            CalcPhaseScope phase
            CalcPhaseScope phase
        Case Else
            Application.CalculateFull
    End Select

    ' Phase telemetry — accumulate per-project so we can size the speedup
    ' from real portfolio runs and shrink scopes further if data supports it.
    Dim dElapsed As Double
    dElapsed = Timer - t0
    Select Case phase
        Case PHASE_DSCR
            mCalcSecsDSCR = mCalcSecsDSCR + dElapsed
        Case PHASE_NPP
            mCalcSecsNPP = mCalcSecsNPP + dElapsed
        Case PHASE_APPR
            mCalcSecsAppr = mCalcSecsAppr + dElapsed
        Case Else
            mCalcSecsFull = mCalcSecsFull + dElapsed
    End Select

    DoEvents  ' Yield to COM message pump during long solve loops
End Sub

Private Sub CalcPhaseScope(ByVal phase As Integer)
    ' Run the per-sheet recalc set appropriate for `phase`. The full-scope
    ' fallback (CalcSheetsAll) is taken whenever phase is unknown or the
    ' tight-scope toggle for that phase is disabled inside the helper.
    Select Case phase
        Case PHASE_DSCR
            CalcSheetsForDSCR
        Case PHASE_NPP
            CalcSheetsForNPP
        Case PHASE_APPR
            CalcSheetsForAppraisal
        Case Else
            CalcSheetsAll
    End Select
End Sub

Private Sub CalcSheetsAll()
    ' Full 13-sheet recalc — original CalcCoreSheetsHL, unchanged.
    Sheets("Project Inputs").Calculate
    Sheets("Rate Curves").Calculate
    Sheets("Ops Sandbox").Calculate
    Sheets("Global").Calculate
    Sheets("Operations").Calculate
    Sheets("Capex").Calculate
    Sheets("Safe Harbor").Calculate
    Sheets("CL").Calculate
    Sheets("Perm Debt").Calculate
    Sheets("Tax Equity").Calculate
    Sheets("Appraisal").Calculate
    Sheets("NPP Calc").Calculate
    Sheets("PT Returns").Calculate
End Sub

Private Sub CalcSheetsForDSCR()
    ' DSCR GoalSeek changes PT Returns!F129. Rate Curves / Ops Sandbox /
    ' Global / Capex / Safe Harbor are upstream of cap-structure inputs and
    ' do not depend on DSCR. Project Inputs reads from PT Returns but doesn't
    ' feed back into the DSCR-dependent chain within a single GoalSeek step.
    If Not USE_TIGHT_DSCR_SCOPE Then
        CalcSheetsAll
        Exit Sub
    End If
    Sheets("Operations").Calculate
    Sheets("CL").Calculate
    Sheets("Perm Debt").Calculate
    Sheets("Tax Equity").Calculate
    Sheets("PT Returns").Calculate
End Sub

Private Sub CalcSheetsForNPP()
    ' NPP cell change on Project Inputs ripples through nearly the whole
    ' downstream chain (appraised value, capital stack, returns). Default
    ' to full scope; flipping USE_TIGHT_NPP_SCOPE = True drops only the
    ' upstream-only sheets (Rate Curves, Ops Sandbox, Global, Capex, Safe
    ' Harbor) once that's been validated against a known portfolio.
    If Not USE_TIGHT_NPP_SCOPE Then
        CalcSheetsAll
        Exit Sub
    End If
    Sheets("Project Inputs").Calculate
    Sheets("Operations").Calculate
    Sheets("CL").Calculate
    Sheets("Perm Debt").Calculate
    Sheets("Tax Equity").Calculate
    Sheets("Appraisal").Calculate
    Sheets("NPP Calc").Calculate
    Sheets("PT Returns").Calculate
End Sub

Private Sub CalcSheetsForAppraisal()
    ' Dev Fee change on Project Inputs feeds the Appraisal sheet's valuation,
    ' which feeds back to Project Inputs!F31 (rApprLive). The Appraisal IRR
    ' is a pre-tax valuation construct and does not require PT Returns to
    ' reconverge between GoalSeek attempts within a single inner loop.
    If Not USE_TIGHT_APPR_SCOPE Then
        CalcSheetsAll
        Exit Sub
    End If
    Sheets("Project Inputs").Calculate
    Sheets("Appraisal").Calculate
    Sheets("NPP Calc").Calculate
End Sub

Private Sub ResetPhaseTelemetryHL()
    mCalcSecsDSCR = 0
    mCalcSecsNPP = 0
    mCalcSecsAppr = 0
    mCalcSecsFull = 0
End Sub

' Per-project elapsed seconds, robust to VBA Timer's midnight rollover.
' Timer returns seconds since midnight (0..86400); subtraction goes
' negative when the solve crosses 00:00:00. Without correction, the
' per-project timeout check below fails and a runaway project can grind
' for hours. Add 86400 once if Timer wrapped (single-day budgets only).
Private Function ProjectElapsedHL(ByVal dStart As Double) As Double
    Dim dEl As Double
    dEl = Timer - dStart
    If dEl < 0 Then dEl = dEl + 86400
    ProjectElapsedHL = dEl
End Function

Private Sub ResetCalcTierHL()
    mCalcTier = 1
End Sub

Private Sub EscalateCalcTierHL()
    If mCalcTier < 3 Then mCalcTier = mCalcTier + 1
End Sub

Private Sub CalcOutputSheetsHL()
    ' Output sheets split into two groups so the skip flag never starves
    ' critical downstream tabs:
    '
    '   ALWAYS recalc: Dashboard, Table
    '     -- summary surfaces the deal team relies on every run; per
    '        Caroline's spec these must reflect the converged state.
    '
    '   SKIP-IF-FLAGGED: Portfolio, AT Returns_WIP, Corp Model Output,
    '                    Cust Prop, Waterfall Sensitivity
    '     -- portfolio rollups and scenario tabs that the team typically
    '        inspects separately. Stale state here is acceptable for
    '        speed; Excel recalcs lazily on next interactive open.
    '
    ' mSkipOutputRecalc=True only affects the second group. Set via
    ' SetSkipOutputRecalcHL from Python before Init / SolveHeadless.

    Dim vAlways As Variant
    Dim vOptional As Variant
    Dim vSh As Variant

    vAlways = Array("Dashboard", "Table")
    For Each vSh In vAlways
        On Error Resume Next
        Sheets(CStr(vSh)).EnableCalculation = True
        Sheets(CStr(vSh)).Calculate
        On Error GoTo 0
    Next

    If mSkipOutputRecalc Then Exit Sub

    vOptional = Array("Portfolio", "AT Returns_WIP", "Corp Model Output", _
                      "Cust Prop", "Waterfall Sensitivity")
    For Each vSh In vOptional
        On Error Resume Next
        Sheets(CStr(vSh)).EnableCalculation = True
        Sheets(CStr(vSh)).Calculate
        On Error GoTo 0
    Next
End Sub

Public Sub SetSkipOutputRecalcHL(ByVal bSkip As Boolean)
    ' Public setter for the output-sheet recalc toggle.
    ' Python calls this via Application.Run before InitSolveEnvHL or
    ' single-shot SolveHeadless. Both code paths funnel through
    ' CalcOutputSheetsHL, so guarding there covers chunked and non-chunked
    ' runs with one switch.
    mSkipOutputRecalc = bSkip
End Sub

Private Sub DisableNonCoreSheets()
    Dim vCore As Variant
    vCore = Array("Project Inputs", "Rate Curves", "Ops Sandbox", "Global", _
                  "Operations", "Capex", "Safe Harbor", "CL", _
                  "Perm Debt", "Tax Equity", "Appraisal", "NPP Calc", "PT Returns")

    Dim ws     As Worksheet
    Dim bCore  As Boolean
    Dim vSh    As Variant

    For Each ws In ThisWorkbook.Worksheets
        bCore = False
        For Each vSh In vCore
            If ws.Name = CStr(vSh) Then
                bCore = True
                Exit For
            End If
        Next
        If Not bCore Then
            On Error Resume Next
            ws.EnableCalculation = False
            On Error GoTo 0
        End If
    Next
End Sub

Private Sub EnableAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.EnableCalculation = True
        On Error GoTo 0
    Next
End Sub

Private Sub SetGoalSeekPrecisionHL()
    ' Match working SolveMinEquityWithHoldCo precision. Looser settings cause
    ' Appraisal GoalSeek to exit before Dev Fee has actually converged.
    Application.MaxChange = 0.00001
    Application.MaxIterations = GS_MAXITER_COLD
End Sub

Private Sub SetGoalSeekModeHL(ByVal bColdMode As Boolean)
    If bColdMode Then
        Application.MaxIterations = GS_MAXITER_COLD
    Else
        Application.MaxIterations = GS_MAXITER_WARM
    End If
End Sub

Private Sub RestoreGoalSeekDefaultsHL()
    Application.MaxChange = 0.001
    Application.MaxIterations = 100
End Sub


' ==============================================================================
'  PUBLIC: Targeted recalc for Python result-reading
' ==============================================================================
Public Sub SwitchProjectAndRecalc(ByVal projOffset As Integer)
    ThisWorkbook.Sheets(SHT_PI).Range(PI_PROJ_INDEX).Value = projOffset
    ResetCalcTierHL
    CalcModelCoreHL
End Sub


' ==============================================================================
'  PUBLIC: Snapshot the active project's per-column convergence cells as
'  hard constants. Called by Python from the post-solve read pass, AFTER
'  SwitchProjectAndRecalc has set F2 = this project's offset and refreshed
'  the workbook state.
'
'  Stamps rows 31, 32, 33, 37, 38, 39 on Project Inputs at column `colIdx`,
'  replacing each cell's IF(<col>2=$F$2, $F$<row>, <col><row>) circular-
'  formula cache with the cell's current evaluated value.
'
'  Why not stamp inside SolveOneProjectByColHL: doing so captures a
'  transient cross-project state because each project's solve modifies
'  cells that NPP Calc!H453 (= F37) and Appraisal cells (= F31, F33)
'  depend on. The "real" converged value for project N can only be read
'  after ALL projects have solved AND F2 is set back to N. The worker's
'  post-read pass does exactly that — stamping there pins the value the
'  worker is about to report in solved_values, and the merged file's
'  per-column cells then match the verifier's expectations.
'
'  Rows 31 (Live Appraisal IRR) and 37 (Live IRR) read from F31 / F37
'  rather than self-assign because the per-column IF formula is a
'  circular reference; reading the per-column cell would resolve to
'  the cached side of the IF (= itself) instead of the F<row> side.
'  Rows 32, 33, 38, 39 self-assign via HardStampNumericHL, which refuses
'  to freeze Excel error values into the cell.
' ==============================================================================
Public Sub StampActiveProjectColumnHL(ByVal colIdx As Integer)
    If colIdx < PI_BASE_COL + 1 Or colIdx > PI_BASE_COL + COL_SCAN_LIMIT Then
        Err.Raise 5, "StampActiveProjectColumnHL", _
            "colIdx " & colIdx & " out of bounds [" & (PI_BASE_COL + 1) & _
            ".." & (PI_BASE_COL + COL_SCAN_LIMIT) & "]"
    End If

    Dim wsPI As Worksheet
    Set wsPI = ThisWorkbook.Sheets(SHT_PI)

    ' Force full propagation BEFORE reading row 31/37. The caller's
    ' SwitchProjectAndRecalc uses tier-1 CalcSheetsAll which doesn't
    ' reliably mark cross-sheet OFFSET-via-F2 dependencies dirty under
    ' xlCalculationManual + multi-threaded calc (same reason the solve
    ' loop runs CalculateFull at line 821 before declaring convergence).
    ' Without this, the hard-stamp could pin a stale prior-project IRR
    ' value into this column — invisible until the merged file diverges
    ' from the worker's reported solved_values.
    '
    ' Lower MaxIterations to 100 around this call. SetGoalSeekPrecisionHL
    ' raised it to 1000 (GS_MAXITER_COLD) for the solve loop. At 1000
    ' iterations a CalculateFull on a workbook with circular sticky-IF
    ' cells in rows 31/37 + #DIV/0 propagation from a partially-populated
    ' column can spin for minutes per call. Read-pass doesn't need that
    ' resolution -- 100 iter is plenty for steady-state propagation.
    Dim lSavedMaxIter As Long
    lSavedMaxIter = Application.MaxIterations
    Application.MaxIterations = 100
    Application.CalculateFull
    Application.MaxIterations = lSavedMaxIter

    ' Live IRR / Live Appraisal — read from F-column live cells, not self-
    ' assign (the per-column IF is circular; self-assign returns the
    ' cached side of the IF).
    wsPI.Cells(37, colIdx).Value = wsPI.Range(PI_IRR_LIVE).Value
    wsPI.Cells(31, colIdx).Value = wsPI.Range(PI_APPR_LIVE).Value

    ' Dev Fee, NPP $/W (GoalSeek changing-cells, already numeric) and
    ' FMV, NPP $ total (formulas) — self-assign via the IsError-guarded
    ' helper.
    HardStampNumericHL wsPI, 32, colIdx, "POSTREAD"
    HardStampNumericHL wsPI, 38, colIdx, "POSTREAD"
    HardStampNumericHL wsPI, 33, colIdx, "POSTREAD"
    HardStampNumericHL wsPI, 39, colIdx, "POSTREAD"
End Sub


' ==============================================================================
'  PUBLIC: VBA-side merge helper for the parallel runner's fallback path.
'
'  When openpyxl-based merge (keep_vba=True round-trip) fails to preserve
'  the workbook's macros or formulas, parallel_runner falls back to opening
'  the master workbook in a fresh Excel session and calling this Sub once
'  per other-worker-owned project, stamping the converged column outputs
'  directly via Excel COM. Excel handles the .xlsm round-trip natively, so
'  this is the durable fallback.
'
'  Stamps the same six rows that VBA hardcoded at end-of-solve into the
'  per-project columns (see SolveOneProjectByColHL / SolveHeadless body).
' ==============================================================================
Public Sub StampConvergedValuesHL(ByVal colIdx As Integer, _
                                   ByVal npp As Double, _
                                   ByVal devFee As Double, _
                                   ByVal fmv As Double, _
                                   ByVal liveIRR As Double, _
                                   ByVal apprLive As Double, _
                                   ByVal nppTotal As Double)
    ' Bounds guardrail. Reject anything outside the project-column band
    ' [PI_BASE_COL+1 .. PI_BASE_COL+COL_SCAN_LIMIT] so a corrupted filename
    ' parse in parallel_runner's _merge_via_vba_fallback can't silently
    ' overwrite the wrong cells. Raise Err.Number=5 (invalid procedure
    ' call) instead of Exit Sub so the parent's Application.Run sees the
    ' COM exception — a silent skip would let the merge "succeed" while
    ' actually leaving the master worker's pre-solve values in place for
    ' that project, which is exactly the silent-corruption mode the merge
    ' verification gate is designed to catch.
    If colIdx < PI_BASE_COL + 1 Or colIdx > PI_BASE_COL + COL_SCAN_LIMIT Then
        Err.Raise 5, "StampConvergedValuesHL", _
            "colIdx " & colIdx & " out of bounds [" & (PI_BASE_COL + 1) & _
            ".." & (PI_BASE_COL + COL_SCAN_LIMIT) & "]"
    End If

    Dim wsPI As Worksheet
    Set wsPI = ThisWorkbook.Sheets(SHT_PI)
    wsPI.Cells(PI_ROW_NPP, colIdx).Value = npp
    wsPI.Cells(PI_ROW_DEV_FEE, colIdx).Value = devFee
    wsPI.Cells(33, colIdx).Value = fmv
    wsPI.Cells(37, colIdx).Value = liveIRR
    wsPI.Cells(31, colIdx).Value = apprLive
    wsPI.Cells(39, colIdx).Value = nppTotal

    ' Append an audit row to __SolverResults so a bad merge is forensically
    ' recoverable. Columns mirror the per-solve schema, prefixed "merge"
    ' in column B to distinguish from solve rows.
    Dim wsRes As Worksheet
    Set wsRes = EnsureSolverResultsSheetHL()
    Dim r As Long
    r = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row + 1
    wsRes.Cells(r, 1).Value = colIdx - PI_BASE_COL
    wsRes.Cells(r, 2).Value = "merge:col=" & colIdx
    wsRes.Cells(r, 3).Value = npp
    wsRes.Cells(r, 4).Value = devFee
    wsRes.Cells(r, 5).Value = fmv
    wsRes.Cells(r, 6).Value = liveIRR
    wsRes.Cells(r, 7).Value = apprLive
    wsRes.Cells(r, 8).Value = nppTotal
    wsRes.Cells(r, 14).Value = CStr(Now)
End Sub


' ==============================================================================
'  PUBLIC: Chunked entry points for Python-driven progress reporting
'
'  Python calls: InitSolveEnvHL -> (SolveOneProjectByColHL per project) -> FinalizeSolveEnvHL
'  Between calls, Python reads the __SolverResults sheet row for status.
'  This keeps the fast in-process solve while restoring per-project progress.
' ==============================================================================

Public Sub InitSolveEnvHL()
    Dim wsRes As Worksheet
    Set wsRes = EnsureSolverResultsSheetHL()
    ResetSolverResultsHL wsRes

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ' Force iterative calc on regardless of workbook XML state. The IL TEST
    ' 2026-05-13 regression came from a workbook saved with iterateDelta
    ' missing from <calcPr/>; we don't want a future regression on the
    ' Iteration flag itself to silently break the row 31 self-circular
    ' per-column hard-stamps. Belt-and-suspenders: workbook says iterate=
    ' True, AND macro asserts it before any solve runs. SetGoalSeekPrecision
    ' below pins MaxChange to 0.00001 (tighter than the 0.0001 used by
    ' the working SMP / RP Puma models).
    Application.Iteration = True
    SetGoalSeekPrecisionHL
    ResetCalcTierHL
    DisableNonCoreSheets

    On Error Resume Next
    Application.MultiThreadedCalculation.Enabled = True
    On Error GoTo 0

    WriteHeartbeatHL wsRes, "INIT|" & CStr(Now)
End Sub

Public Sub FinalizeSolveEnvHL(ByVal lngOriginalF2 As Long)
    Dim wsPI As Worksheet
    Dim wsRes As Worksheet
    Set wsPI = ThisWorkbook.Sheets(SHT_PI)
    Set wsRes = EnsureSolverResultsSheetHL()

    wsPI.Range(PI_PROJ_INDEX).Value = lngOriginalF2
    ResetCalcTierHL
    CalcModelCoreHL
    EnableAllSheets
    CalcOutputSheetsHL

    RestoreGoalSeekDefaultsHL
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    WriteHeartbeatHL wsRes, "COMPLETE|" & CStr(Now)
End Sub

' Solve a single project. Writes results to __SolverResults row `resultsRow`.
' Returns 1 if converged, 0 otherwise (as a Long, since Python COM sees Long cleanly).
Public Function SolveOneProjectByColHL(ByVal colIdx As Integer, _
                                        ByVal projName As String, _
                                        ByVal resultsRow As Integer) As Long
    On Error GoTo ProjErr
    Dim wsPI As Worksheet
    Dim wsPT As Worksheet
    Dim wsRes As Worksheet
    Set wsPI = ThisWorkbook.Sheets(SHT_PI)
    Set wsPT = ThisWorkbook.Sheets(SHT_PT)
    Set wsRes = EnsureSolverResultsSheetHL()

    Dim projOffset As Integer
    projOffset = colIdx - PI_BASE_COL

    Dim dSolveStart As Double
    dSolveStart = Timer
    WriteHeartbeatHL wsRes, "RUNNING|" & CStr(Now) & "|Project=" & projName & "|col=" & colIdx

    ' Defensive initialization of lOriginalHoldCo BEFORE the early-skip
    ' block below. If any line between here and the real HoldCo read
    ' (further down) raises, ProjErr fires with lOriginalHoldCo
    ' uninitialized — VBA Long defaults to 0, which is invalid HoldCo
    ' state and would write through to PT_HOLDCO_ONOFF on the error-
    ' restore path. Default to 1 (HoldCo ON, the post-solve steady
    ' state in every documented TE structure). The "real" HoldCo read
    ' further down overwrites this when reached; if the function exits
    ' via early-skip, no restore happens anyway because no mutation
    ' was attempted.
    Dim lOriginalHoldCo As Long
    lOriginalHoldCo = 1

    ' Pre-flight skip BEFORE the expensive F2 write + CalcForPhase PHASE_FULL
    ' below. The Tranche 7.1 fix moved the same checks here from after the
    ' recalc, but those late checks still paid the full pre-loop calc cost
    ' (5-30 min per placeholder on a 13MB workbook with iterative calc).
    '
    ' 2026-05-18 SMP run id=17 post-mortem: Worker 1 skipped its 5
    ' placeholders fast (iterative calc happened to settle quickly on its
    ' first hit); Worker 0 stalled 28+ min on Project 8 — iterative calc
    ' on the first placeholder it encountered (after 3 real solves left
    ' state in a non-converging configuration) burned the full per-call
    ' watchdog window. Asymmetric failure between identical workers on
    ' byte-identical inputs proved the late placement was unsafe.
    '
    ' The skip metadata reads come from literal-input cells: row 11
    ' (MWdc), row 159 (RC1 toggle), row 160 (RC1 Generic rate at COD).
    ' These are populated directly by the operator / portfolio mapper —
    ' never formulas — so reading them by direct column index before any
    ' recalc is safe and deterministic.
    Dim dEarlyMWdc As Double
    dEarlyMWdc = 0
    On Error Resume Next
    dEarlyMWdc = CDbl(wsPI.Cells(11, colIdx).Value)
    On Error GoTo ProjErr
    If dEarlyMWdc <= 0 Then
        wsRes.Cells(resultsRow, 1).Value = projOffset
        wsRes.Cells(resultsRow, 2).Value = projName
        wsRes.Cells(resultsRow, 9).Value = False
        wsRes.Cells(resultsRow, 12).Value = "skipped:no_mwdc"
        wsRes.Cells(resultsRow, 13).Value = Round(ProjectElapsedHL(dSolveStart), 4)
        wsRes.Cells(resultsRow, 14).Value = CStr(Now)
        wsRes.Cells(resultsRow, 20).Value = "skipped"
        WriteHeartbeatHL wsRes, "SKIPPED|" & projName & "|MWdc=0_or_invalid"
        SolveOneProjectByColHL = 0
        Exit Function
    End If

    ' Custom toggle path is NOT caught here — Custom rates source from
    ' the Rate Curves tab vector, so a zero in row 160 means nothing for
    ' them. E15a/E16 preflight catches Custom-with-empty-Name separately.
    Dim vEarlyRC1Tog As Variant
    Dim vEarlyRC1Rate As Variant
    Dim bEarlyRC1Empty As Boolean
    bEarlyRC1Empty = False
    vEarlyRC1Tog = wsPI.Cells(159, colIdx).Value
    vEarlyRC1Rate = wsPI.Cells(160, colIdx).Value
    If CStr(vEarlyRC1Tog) = "Generic" Then
        If IsNumeric(vEarlyRC1Rate) Then
            If CDbl(vEarlyRC1Rate) = 0 Then bEarlyRC1Empty = True
        ElseIf IsEmpty(vEarlyRC1Rate) Or IsNull(vEarlyRC1Rate) Then
            bEarlyRC1Empty = True
        End If
    End If
    If bEarlyRC1Empty Then
        wsRes.Cells(resultsRow, 1).Value = projOffset
        wsRes.Cells(resultsRow, 2).Value = projName
        wsRes.Cells(resultsRow, 9).Value = False
        wsRes.Cells(resultsRow, 12).Value = "skipped:no_rc1_revenue"
        wsRes.Cells(resultsRow, 13).Value = Round(ProjectElapsedHL(dSolveStart), 4)
        wsRes.Cells(resultsRow, 14).Value = CStr(Now)
        wsRes.Cells(resultsRow, 20).Value = "skipped"
        WriteHeartbeatHL wsRes, "SKIPPED|" & projName & "|RC1_Generic_zero_rate"
        SolveOneProjectByColHL = 0
        Exit Function
    End If

    ' Save HoldCo state at entry so the error path can restore it. The
    ' iter loop toggles HoldCo 0→1 within each iIter; an error fired with
    ' HoldCo=0 would leave the workbook in an inconsistent capital
    ' structure state for subsequent project solves until the next iter's
    ' HoldCo cycle. Restoring on error eliminates that cross-project
    ' bleed (P0-1 from 2026-05-15 reliability review).
    '
    ' GUARD: PT!C134 (the HoldCo toggle) is a formula that depends on the
    ' currently-active F2. At this point F2 still points at the PREVIOUS
    ' project, and on rare occasions C134 resolves to an error token
    ' (#DIV/0! / #VALUE!) when the previous project left the cap stack
    ' in a transient state. A bare CLng() on an error value raises VBA
    ' Type Mismatch (13) → 0x800a9c68 to COM, crashing the worker.
    ' Default to 1 (HoldCo ON) on any non-numeric read — the value is
    ' only used for restore-on-error and HoldCo=1 is the post-solve
    ' steady state in every documented TE structure. (2026-05-18 SMP
    ' post-mortem.)
    ' lOriginalHoldCo was defensively initialized to 1 at function entry
    ' (Tranche 7.7); reassign here from the live cell when we reach this
    ' point (post-skip, F2 is about to be written and the real
    ' restore-on-error path is meaningful).
    Dim vHoldCoRaw As Variant
    vHoldCoRaw = wsPT.Range(PT_HOLDCO_ONOFF).Value
    If IsNumeric(vHoldCoRaw) Then
        lOriginalHoldCo = CLng(vHoldCoRaw)
        If lOriginalHoldCo <> 0 And lOriginalHoldCo <> 1 Then lOriginalHoldCo = 1
    Else
        lOriginalHoldCo = 1
    End If

    ' Per-step heartbeats. The cost is ~20 extra cell writes per project
    ' (negligible vs the 60-200s solve time) and they pay for themselves
    ' the first time a workbook fails — the last heartbeat in
    ' __SolverResults!N1 names the exact line that threw. Promoted from
    ' debug instrumentation on 2026-05-13 after they pinpointed an
    ' openpyxl-save corruption issue on the RP Puma 2026.05.08 workbooks.
    ' Route OFFSET formulas to this project
    WriteHeartbeatHL wsRes, "STEP1_F2_WRITE|" & projName
    wsPI.Range(PI_PROJ_INDEX).Value = projOffset
    WriteHeartbeatHL wsRes, "STEP2_RESET_CALC_TIER"
    ResetCalcTierHL
    WriteHeartbeatHL wsRes, "STEP3_RESET_PHASE_TELEMETRY"
    ResetPhaseTelemetryHL
    WriteHeartbeatHL wsRes, "STEP4_CALC_FULL_PRE_LOOP"
    CalcForPhase PHASE_FULL

    WriteHeartbeatHL wsRes, "STEP5_RANGE_SETUP"

    Dim rHoldCo     As Range
    Dim rEquity     As Range
    Dim rMinEqTgt   As Range
    Dim rDSCR       As Range
    Dim rTotalUses  As Range
    Dim rIRRLive    As Range
    Dim rIRRTgt     As Range
    Dim rNPP        As Range
    Dim rApprLive   As Range
    Dim rWACCTgt    As Range
    Dim rDevFee     As Range

    Set rHoldCo = wsPT.Range(PT_HOLDCO_ONOFF)
    Set rEquity = wsPT.Range(PT_EQUITY)
    Set rMinEqTgt = wsPT.Range(PT_MIN_EQ_TARGET)
    Set rDSCR = wsPT.Range(PT_DSCR_MULTIPLE)
    Set rTotalUses = wsPT.Range(PT_TOTAL_USES)
    Set rIRRLive = wsPI.Range(PI_IRR_LIVE)
    Set rIRRTgt = wsPI.Range(PI_IRR_TARGET)
    Set rNPP = wsPI.Cells(PI_ROW_NPP, colIdx)
    Set rApprLive = wsPI.Range(PI_APPR_LIVE)
    Set rWACCTgt = wsPI.Range(PI_WACC_TARGET)
    Set rDevFee = wsPI.Cells(PI_ROW_DEV_FEE, colIdx)
    WriteHeartbeatHL wsRes, "STEP6_PRE_SEED"

    ' Pre-seed only if cell is blank. Do NOT clamp by bounds --
    ' SolveMinEquityWithHoldCo trusts GoalSeek, and clamping legitimate
    ' answers to the seed prevents Appraisal from ever converging.
    If rNPP.Value = "" Then rNPP.Value = NPP_SEED
    If rDevFee.Value = "" Then rDevFee.Value = DEV_FEE_SEED
    WriteHeartbeatHL wsRes, "STEP7_PRE_LOOP_INIT"

    Dim bConverged  As Boolean
    Dim bColdMode   As Boolean
    Dim bNeedCold   As Boolean
    Dim iGSRetry    As Integer
    Dim iIter       As Integer
    Dim iInner      As Integer
    Dim iActualIters As Integer
    Dim bGSok       As Boolean
    Dim dEquityPct  As Double
    Dim dPrevEqPct  As Double
    Dim dIRRGap     As Double
    Dim dApprGap    As Double
    Dim dTotalUses  As Double
    Dim dPrevNPP    As Double
    Dim dPrevDevFee As Double

    bConverged = False
    bColdMode = False
    iGSRetry = MAX_GS_RETRY_WARM
    SetGoalSeekModeHL False
    ResetCalcTierHL
    dPrevEqPct = -999
    iActualIters = 0

    For iIter = 1 To MAX_ITER
        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_START"
        If ProjectElapsedHL(dSolveStart) > PROJECT_TIMEOUT_SECONDS Then
            iActualIters = iIter - 1
            Exit For
        End If
        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_HOLDCO_OFF"
        rHoldCo.Value = 0
        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_CALC1"
        CalcForPhase PHASE_FULL

        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_GS_DSCR"
        ' GoalSeek RAISES (not just returns False) when the dependent formula
        ' chain produces an error value — e.g. PT_EQUITY = #DIV/0! when a
        ' placeholder project has zero revenue. The bGSok = ... assignment
        ' never completes in that case; the raise jumps to ProjErr and the
        ' COM worker crashes with 0x800a9c68. Localize the OERN so a real
        ' VBA error (not GoalSeek's failure mode) still surfaces.
        On Error Resume Next
        bGSok = rEquity.GoalSeek(Goal:=rMinEqTgt.Value, ChangingCell:=rDSCR)
        If Err.Number <> 0 Then
            bGSok = False
            WriteHeartbeatHL wsRes, "GS_RAISE_DSCR|iter=" & iIter & "|err=" & Err.Number
            Err.Clear
        End If
        On Error GoTo ProjErr
        If Not bGSok Then WriteHeartbeatHL wsRes, "GS_FAIL_DSCR|iter=" & iIter
        If rDSCR.Value < DSCR_MIN Then rDSCR.Value = DSCR_MIN
        If rDSCR.Value > DSCR_MAX Then rDSCR.Value = DSCR_MAX
        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_CALC2"
        CalcForPhase PHASE_DSCR

        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_HOLDCO_ON"
        rHoldCo.Value = 1
        WriteHeartbeatHL wsRes, "ITER_" & iIter & "_CALC3"
        CalcForPhase PHASE_FULL

        dPrevNPP = -999#
        dPrevDevFee = -999#

        For iInner = 1 To iGSRetry
            If ProjectElapsedHL(dSolveStart) > PROJECT_TIMEOUT_SECONDS Then Exit For

            WriteHeartbeatHL wsRes, "ITER_" & iIter & "_INNER_" & iInner & "_GS_NPP"
            On Error Resume Next
            bGSok = rIRRLive.GoalSeek(Goal:=rIRRTgt.Value, ChangingCell:=rNPP)
            If Err.Number <> 0 Then
                bGSok = False
                WriteHeartbeatHL wsRes, "GS_RAISE_NPP|iter=" & iIter & "|inner=" & iInner & "|err=" & Err.Number
                Err.Clear
            End If
            On Error GoTo ProjErr
            If Not bGSok Then WriteHeartbeatHL wsRes, "GS_FAIL_NPP|iter=" & iIter & "|inner=" & iInner
            WriteHeartbeatHL wsRes, "ITER_" & iIter & "_INNER_" & iInner & "_CALC_NPP"
            CalcForPhase PHASE_NPP

            WriteHeartbeatHL wsRes, "ITER_" & iIter & "_INNER_" & iInner & "_GS_APPR"
            On Error Resume Next
            bGSok = rApprLive.GoalSeek(Goal:=rWACCTgt.Value, ChangingCell:=rDevFee)
            If Err.Number <> 0 Then
                bGSok = False
                WriteHeartbeatHL wsRes, "GS_RAISE_APPR|iter=" & iIter & "|inner=" & iInner & "|err=" & Err.Number
                Err.Clear
            End If
            On Error GoTo ProjErr
            If Not bGSok Then WriteHeartbeatHL wsRes, "GS_FAIL_APPR|iter=" & iIter & "|inner=" & iInner
            WriteHeartbeatHL wsRes, "ITER_" & iIter & "_INNER_" & iInner & "_CALC_APPR"
            CalcForPhase PHASE_APPR

            ' Defensive Abs(numeric - numeric): if GoalSeek left rIRRLive
            ' or rApprLive in an error state (#DIV/0 from broken revenue),
            ' arithmetic raises Type Mismatch. Treat non-numeric as a large
            ' gap so the loop continues toward Exit For via the slope-stall
            ' break rather than crashing the whole project.
            If IsNumeric(rIRRLive.Value) And IsNumeric(rIRRTgt.Value) Then
                dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
            Else
                dIRRGap = 1#
            End If
            If IsNumeric(rApprLive.Value) And IsNumeric(rWACCTgt.Value) Then
                dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
            Else
                dApprGap = 1#
            End If
            If dIRRGap <= IRR_TOLERANCE And dApprGap <= APPR_TOLERANCE Then
                ' Phase scopes can leave F37 reading a stale NPP Calc!H453.
                ' Validate against full propagation before declaring conv-
                ' ergence — Application.CalculateFull is the only call that
                ' reliably re-evaluates cross-sheet OFFSET-via-F2 chains.
                Application.CalculateFull
                If IsNumeric(rIRRLive.Value) And IsNumeric(rIRRTgt.Value) Then
                    dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
                Else
                    dIRRGap = 1#
                End If
                If IsNumeric(rApprLive.Value) And IsNumeric(rWACCTgt.Value) Then
                    dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
                Else
                    dApprGap = 1#
                End If
                If dIRRGap <= IRR_TOLERANCE And dApprGap <= APPR_TOLERANCE Then Exit For
            End If

            ' Slope-stall break: GoalSeek made no measurable progress on
            ' either changing cell vs. the prior retry. Further retries at
            ' this calc tier won't move the answer — escalate at the outer
            ' level instead.
            If iInner > 1 _
               And Abs(rNPP.Value - dPrevNPP) < 0.000001 _
               And Abs(rDevFee.Value - dPrevDevFee) < 0.000001 Then
                Exit For
            End If
            dPrevNPP = rNPP.Value
            dPrevDevFee = rDevFee.Value

            EscalateCalcTierHL
        Next iInner

        ' Defensive read: a broken cap stack (placeholder project with zero
        ' revenue) leaves rTotalUses / rEquity as #DIV/0! and the bare
        ' arithmetic below would raise Type Mismatch. Treat non-numeric as
        ' equityPct=0 so the outer convergence check naturally fails and
        ' we continue to the result-write path with bConverged=False.
        If IsNumeric(rTotalUses.Value) And IsNumeric(rEquity.Value) Then
            dTotalUses = rTotalUses.Value
            If dTotalUses <> 0 Then
                dEquityPct = rEquity.Value / dTotalUses
            Else
                dEquityPct = 0
            End If
        Else
            dTotalUses = 0
            dEquityPct = 0
        End If

        If Abs(dEquityPct - 0.1) <= EQUITY_FINAL_TOL Then
            bConverged = True
            iActualIters = iIter
            Exit For
        End If
        ' Relaxed-tier early exit: equity within +/-0.5pp and inner gaps
        ' within 5x tolerance is investment-grade. Skip the remaining
        ' outer iterations -- they rarely tighten further. bConverged
        ' stays False so column I keeps strict-only semantics; column T
        ' carries the tier for Python-side policy.
        If Abs(dEquityPct - 0.1) <= EQUITY_RELAXED_TOL _
           And dIRRGap  <= IRR_TOLERANCE  * RELAXED_GAP_FACTOR _
           And dApprGap <= APPR_TOLERANCE * RELAXED_GAP_FACTOR Then
            iActualIters = iIter
            Exit For
        End If
        If Abs(dEquityPct - dPrevEqPct) < 0.000005 And iIter > 1 Then
            iActualIters = iIter
            Exit For
        End If

        ' Soft-cap exit: once we've burned MAX_ITER_SOFT iterations and the
        ' last equity move was below 0.1pp, the project has settled --
        ' further passes won't tighten it. The hard cap (MAX_ITER) still
        ' applies to projects that are still making real progress.
        If iIter >= MAX_ITER_SOFT Then
            If Abs(dEquityPct - dPrevEqPct) < SOFT_CAP_PROGRESS_TOL Then
                iActualIters = iIter
                Exit For
            End If
        End If
        dPrevEqPct = dEquityPct

        ' Residual-gated cold-mode escalation. Iter 1's actual residuals
        ' decide whether iter 2+ needs the heavier retries / calc tier;
        ' a portfolio that's already close to converged stays warm.
        bNeedCold = (Abs(dEquityPct - 0.1) > 0.02) _
                    Or (dIRRGap > IRR_TOLERANCE * 3) _
                    Or (dApprGap > APPR_TOLERANCE * 3)

        If Not bColdMode And bNeedCold Then
            bColdMode = True
            iGSRetry = MAX_GS_RETRY_COLD
            SetGoalSeekModeHL True
            EscalateCalcTierHL
        End If
    Next iIter

    ' For-loop variable is one past MAX_ITER on natural completion; on an
    ' Exit For the assignment inside the loop already captured the count.
    If iActualIters = 0 Then iActualIters = MAX_ITER

    ' Force full propagation so dIRRGap / dApprGap below are measured against
    ' the truly-converged F37 / F31, not stale phase-scoped values.
    Application.CalculateFull
    ' Same defensive read as in the inner loop — a project that broke into
    ' an error state still needs to fall through to the result-write path
    ' with bConverged=False rather than crash here.
    If IsNumeric(rIRRLive.Value) And IsNumeric(rIRRTgt.Value) Then
        dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
    Else
        dIRRGap = 1#
    End If
    If IsNumeric(rApprLive.Value) And IsNumeric(rWACCTgt.Value) Then
        dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
    Else
        dApprGap = 1#
    End If

    ' NOTE: per-column hard-stamps for rows 31/32/33/37/38/39 USED to
    ' happen here, but the values they captured were a transient cross-
    ' project state — F37 / F31 / F33 / F39 depend on cell chains that
    ' shift as subsequent projects' solves modify other columns. The
    ' verifier (validation/post_merge.verify_merged_file) caught the
    ' divergence on the SMP WalkTEST 2026-05-12 run. Stamps moved to
    ' StampActiveProjectColumnHL which Python invokes from the post-
    ' solve read pass, AFTER all projects have solved AND F2 is set
    ' back to this project. That's the only moment the workbook is
    ' simultaneously consistent for every project's column.

    wsRes.Cells(resultsRow, 1).Value = projOffset
    wsRes.Cells(resultsRow, 2).Value = projName
    wsRes.Cells(resultsRow, 3).Value = rDSCR.Value
    wsRes.Cells(resultsRow, 4).Value = rNPP.Value
    wsRes.Cells(resultsRow, 5).Value = rDevFee.Value
    wsRes.Cells(resultsRow, 6).Value = dEquityPct
    wsRes.Cells(resultsRow, 7).Value = dIRRGap
    wsRes.Cells(resultsRow, 8).Value = dApprGap
    wsRes.Cells(resultsRow, 9).Value = bConverged
    wsRes.Cells(resultsRow, 10).Value = mCalcTier
    wsRes.Cells(resultsRow, 11).Value = iGSRetry
    wsRes.Cells(resultsRow, 12).Value = IIf(bColdMode, "cold", "warm")
    wsRes.Cells(resultsRow, 13).Value = Round(ProjectElapsedHL(dSolveStart), 4)
    wsRes.Cells(resultsRow, 14).Value = CStr(Now)
    wsRes.Cells(resultsRow, 15).Value = Round(mCalcSecsDSCR, 4)
    wsRes.Cells(resultsRow, 16).Value = Round(mCalcSecsNPP, 4)
    wsRes.Cells(resultsRow, 17).Value = Round(mCalcSecsAppr, 4)
    wsRes.Cells(resultsRow, 18).Value = Round(mCalcSecsFull, 4)
    wsRes.Cells(resultsRow, 19).Value = iActualIters
    wsRes.Cells(resultsRow, 20).Value = ClassifyConvergenceHL(dEquityPct, dIRRGap, dApprGap)
    WriteHeartbeatHL wsRes, "DONE|" & CStr(Now) & "|Project=" & projName

    SolveOneProjectByColHL = IIf(bConverged, 1, 0)
    Exit Function

ProjErr:
    ' Capture error details before any cleanup that could clobber Err.
    Dim lErrNum As Long
    Dim sErrDesc As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    On Error Resume Next
    ' Restore HoldCo so subsequent projects don't inherit a 0 state from
    ' a partial solve. wsPT may be Nothing if Set failed above; guard.
    If Not wsPT Is Nothing Then
        wsPT.Range(PT_HOLDCO_ONOFF).Value = lOriginalHoldCo
    End If
    If Not wsRes Is Nothing Then
        WriteHeartbeatHL wsRes, _
            "PROJ_ERR|" & projName & "|err=" & lErrNum & "|" & sErrDesc
    End If
    On Error GoTo 0
    ' Re-raise so Python's COM caller sees the failure and can decide
    ' whether to retry, recover, or abort the worker.
    Err.Raise lErrNum, _
        "SolveOneProjectByColHL:" & projName, sErrDesc
End Function


' ==============================================================================
'  MAIN ENTRY POINT
' ==============================================================================
Public Sub SolveHeadless()

    ' --- Worksheet references ---
    Dim wsPI As Worksheet
    Dim wsPT As Worksheet
    Dim wsRes As Worksheet
    On Error GoTo ErrHandler
    Set wsPI = ThisWorkbook.Sheets(SHT_PI)
    Set wsPT = ThisWorkbook.Sheets(SHT_PT)
    Set wsRes = EnsureSolverResultsSheetHL()
    ResetSolverResultsHL wsRes

    ' --- Save original F2 for restore ---
    Dim lngOriginalF2 As Long
    lngOriginalF2 = CLng(wsPI.Range(PI_PROJ_INDEX).Value)

    ' --- Save original HoldCo for restore on error ---
    ' Capital structure state must be restored if mid-portfolio error
    ' aborts the loop with HoldCo=0; otherwise the workbook saves with
    ' an inconsistent debt-stack state across projects (P0-1 from review).
    Dim lngOriginalHoldCo As Long
    lngOriginalHoldCo = CLng(wsPT.Range(PT_HOLDCO_ONOFF).Value)
    If lngOriginalHoldCo <> 0 And lngOriginalHoldCo <> 1 Then lngOriginalHoldCo = 1

    ' --- Scan row 7 for toggled-on projects ---
    Dim arrCols(1 To 60)  As Integer
    Dim arrNames(1 To 60) As String
    Dim intOn As Integer
    intOn = 0

    Dim c As Integer
    For c = PI_FIRST_PROJ_COL To PI_FIRST_PROJ_COL + COL_SCAN_LIMIT - 1
        If wsPI.Cells(PI_ROW_NAME, c).Value = "" Then Exit For
        If wsPI.Cells(PI_ROW_TOGGLE, c).Value = 1 Then
            intOn = intOn + 1
            arrCols(intOn) = c
            arrNames(intOn) = CStr(wsPI.Cells(PI_ROW_NAME, c).Value)
        End If
    Next c

    If intOn = 0 Then Exit Sub

    ' --- Performance setup ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ' Force iterative calc on regardless of workbook XML state (see the
    ' matching comment in InitSolveEnvHL for full rationale).
    Application.Iteration = True
    SetGoalSeekPrecisionHL
    ResetCalcTierHL
    DisableNonCoreSheets

    On Error Resume Next
    Application.MultiThreadedCalculation.Enabled = True
    On Error GoTo ErrHandler

    ' --- Declare all solve variables before loop ---
    Dim i           As Integer
    Dim colIdx      As Integer
    Dim projOffset  As Integer
    Dim iIter       As Integer
    Dim iInner      As Integer
    Dim iGSRetry    As Integer
    Dim iActualIters As Integer
    Dim bConverged  As Boolean
    Dim bGSok       As Boolean
    Dim bColdMode   As Boolean
    Dim bNeedCold   As Boolean
    Dim dEquityPct  As Double
    Dim dPrevEqPct  As Double
    Dim dIRRGap     As Double
    Dim dApprGap    As Double
    Dim dTotalUses  As Double
    Dim dSolveStart As Double
    Dim dPrevNPP    As Double
    Dim dPrevDevFee As Double

    Dim rHoldCo     As Range
    Dim rEquity     As Range
    Dim rMinEqTgt   As Range
    Dim rDSCR       As Range
    Dim rTotalUses  As Range
    Dim rIRRLive    As Range
    Dim rIRRTgt     As Range
    Dim rNPP        As Range
    Dim rApprLive   As Range
    Dim rWACCTgt    As Range
    Dim rDevFee     As Range

    ' --- Solve each project ---
    For i = 1 To intOn

        colIdx = arrCols(i)
        projOffset = colIdx - PI_BASE_COL
        dSolveStart = Timer
        WriteHeartbeatHL wsRes, "RUNNING|" & CStr(Now) & "|Project=" & arrNames(i)

        ' Route OFFSET formulas to this project
        wsPI.Range(PI_PROJ_INDEX).Value = projOffset
        ResetPhaseTelemetryHL
        CalcForPhase PHASE_FULL

        ' Set up ranges for this project
        Set rHoldCo = wsPT.Range(PT_HOLDCO_ONOFF)
        Set rEquity = wsPT.Range(PT_EQUITY)
        Set rMinEqTgt = wsPT.Range(PT_MIN_EQ_TARGET)
        Set rDSCR = wsPT.Range(PT_DSCR_MULTIPLE)
        Set rTotalUses = wsPT.Range(PT_TOTAL_USES)
        Set rIRRLive = wsPI.Range(PI_IRR_LIVE)
        Set rIRRTgt = wsPI.Range(PI_IRR_TARGET)
        Set rNPP = wsPI.Cells(PI_ROW_NPP, colIdx)
        Set rApprLive = wsPI.Range(PI_APPR_LIVE)
        Set rWACCTgt = wsPI.Range(PI_WACC_TARGET)
        Set rDevFee = wsPI.Cells(PI_ROW_DEV_FEE, colIdx)

        bConverged = False
        bColdMode = False
        iGSRetry = MAX_GS_RETRY_WARM
        SetGoalSeekModeHL False
        ResetCalcTierHL
        dPrevEqPct = -999
        iActualIters = 0

        ' Pre-seed only if cell is blank. Do NOT clamp by bounds — clamping
        ' legitimate GoalSeek answers prevents Appraisal from converging when
        ' the deal-side Dev Fee genuinely lies outside [DEV_FEE_MIN, DEV_FEE_MAX].
        ' Bounds are sanity checks surfaced by preflight E13/E14, not constraints.
        ' Matches SolveOneProjectByColHL's pattern (lines 747-748).
        If rNPP.Value = "" Then rNPP.Value = NPP_SEED
        If rDevFee.Value = "" Then rDevFee.Value = DEV_FEE_SEED

        For iIter = 1 To MAX_ITER
            If ProjectElapsedHL(dSolveStart) > PROJECT_TIMEOUT_SECONDS Then
                iActualIters = iIter - 1
                Exit For
            End If

            ' Step 1: HoldCo OFF + recalc
            rHoldCo.Value = 0
            CalcForPhase PHASE_FULL

            ' Step 2: GoalSeek Min Equity = 10% (changes DSCR Multiple)
            bGSok = rEquity.GoalSeek(Goal:=rMinEqTgt.Value, ChangingCell:=rDSCR)
            If Not bGSok Then WriteHeartbeatHL wsRes, "GS_FAIL_DSCR|" & arrNames(i) & "|iter=" & iIter
            If rDSCR.Value < DSCR_MIN Then rDSCR.Value = DSCR_MIN
            If rDSCR.Value > DSCR_MAX Then rDSCR.Value = DSCR_MAX
            CalcForPhase PHASE_DSCR

            ' Step 3: HoldCo ON + recalc
            rHoldCo.Value = 1
            CalcForPhase PHASE_FULL

            dPrevNPP = -999#
            dPrevDevFee = -999#

            ' Steps 4-5: Sequential NPP / Appraisal solve
            ' Sequential (not batched) ensures each GoalSeek sees fresh recalc
            ' values — critical for cold-start solves with seed values far from optimal.
            For iInner = 1 To iGSRetry
                If ProjectElapsedHL(dSolveStart) > PROJECT_TIMEOUT_SECONDS Then Exit For

                bGSok = rIRRLive.GoalSeek(Goal:=rIRRTgt.Value, ChangingCell:=rNPP)
                If Not bGSok Then WriteHeartbeatHL wsRes, "GS_FAIL_NPP|" & arrNames(i) & "|iter=" & iIter & "|inner=" & iInner
                CalcForPhase PHASE_NPP

                bGSok = rApprLive.GoalSeek(Goal:=rWACCTgt.Value, ChangingCell:=rDevFee)
                If Not bGSok Then WriteHeartbeatHL wsRes, "GS_FAIL_APPR|" & arrNames(i) & "|iter=" & iIter & "|inner=" & iInner
                CalcForPhase PHASE_APPR

                dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
                dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
                If dIRRGap <= IRR_TOLERANCE And dApprGap <= APPR_TOLERANCE Then
                    ' Phase scopes can leave F37 reading a stale NPP Calc!H453.
                    ' Validate against full propagation before declaring conv-
                    ' ergence.
                    Application.CalculateFull
                    dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
                    dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
                    If dIRRGap <= IRR_TOLERANCE And dApprGap <= APPR_TOLERANCE Then Exit For
                End If

                ' Slope-stall break: GoalSeek made no measurable progress on
                ' either changing cell vs. the prior retry. Further retries at
                ' this calc tier won't move the answer — escalate at the outer
                ' level instead.
                If iInner > 1 _
                   And Abs(rNPP.Value - dPrevNPP) < 0.000001 _
                   And Abs(rDevFee.Value - dPrevDevFee) < 0.000001 Then
                    Exit For
                End If
                dPrevNPP = rNPP.Value
                dPrevDevFee = rDevFee.Value

                EscalateCalcTierHL
                WriteHeartbeatHL wsRes, "RETRY|" & CStr(Now) & "|Project=" & arrNames(i) & "|Inner=" & CStr(iInner)
            Next iInner

            ' Step 6: Convergence check
            dTotalUses = rTotalUses.Value
            If dTotalUses <> 0 Then
                dEquityPct = rEquity.Value / dTotalUses
            Else
                dEquityPct = 0
            End If

            If Abs(dEquityPct - 0.1) <= EQUITY_FINAL_TOL Then
                bConverged = True
                iActualIters = iIter
                Exit For
            End If
            ' Relaxed-tier early exit: equity within +/-0.5pp and inner gaps
            ' within 5x tolerance is investment-grade. Skip the remaining
            ' outer iterations -- they rarely tighten further. bConverged
            ' stays False so column I keeps strict-only semantics; column T
            ' carries the tier for Python-side policy.
            If Abs(dEquityPct - 0.1) <= EQUITY_RELAXED_TOL _
               And dIRRGap  <= IRR_TOLERANCE  * RELAXED_GAP_FACTOR _
               And dApprGap <= APPR_TOLERANCE * RELAXED_GAP_FACTOR Then
                iActualIters = iIter
                Exit For
            End If
            If Abs(dEquityPct - dPrevEqPct) < 0.000005 And iIter > 1 Then
                iActualIters = iIter
                Exit For
            End If

            ' Soft-cap exit: once we've burned MAX_ITER_SOFT iterations and
            ' the last equity move was below 0.1pp, the project has settled
            ' -- further passes won't tighten it. The hard cap (MAX_ITER)
            ' still applies to projects still making real progress.
            If iIter >= MAX_ITER_SOFT Then
                If Abs(dEquityPct - dPrevEqPct) < SOFT_CAP_PROGRESS_TOL Then
                    iActualIters = iIter
                    Exit For
                End If
            End If
            dPrevEqPct = dEquityPct

            ' Residual-gated cold-mode escalation. Iter 1's actual residuals
            ' decide whether iter 2+ needs the heavier retries / calc tier;
            ' a portfolio that's already close to converged stays warm.
            bNeedCold = (Abs(dEquityPct - 0.1) > 0.02) _
                        Or (dIRRGap > IRR_TOLERANCE * 3) _
                        Or (dApprGap > APPR_TOLERANCE * 3)

            If Not bColdMode And bNeedCold Then
                bColdMode = True
                iGSRetry = MAX_GS_RETRY_COLD
                SetGoalSeekModeHL True
                EscalateCalcTierHL
            End If

        Next iIter

        ' For-loop variable is one past MAX_ITER on natural completion; on an
        ' Exit For the assignment inside the loop already captured the count.
        If iActualIters = 0 Then iActualIters = MAX_ITER

        ' Force full propagation so dIRRGap / dApprGap below are measured
        ' against the truly-converged F37 / F31, not stale phase-scoped
        ' values.
        Application.CalculateFull
        dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
        dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)

        ' NOTE: per-column hard-stamps moved to StampActiveProjectColumnHL,
        ' invoked by Python from the post-solve read pass. The in-loop
        ' stamps captured a transient cross-project state that the post-
        ' merge verifier flagged as wrong on the SMP WalkTEST 2026-05-12
        ' run. See the matching comment in SolveOneProjectByColHL for
        ' the full rationale.

        wsRes.Cells(i + 1, 1).Value = projOffset
        wsRes.Cells(i + 1, 2).Value = arrNames(i)
        wsRes.Cells(i + 1, 3).Value = rDSCR.Value
        wsRes.Cells(i + 1, 4).Value = rNPP.Value
        wsRes.Cells(i + 1, 5).Value = rDevFee.Value
        wsRes.Cells(i + 1, 6).Value = dEquityPct
        wsRes.Cells(i + 1, 7).Value = dIRRGap
        wsRes.Cells(i + 1, 8).Value = dApprGap
        wsRes.Cells(i + 1, 9).Value = bConverged
        wsRes.Cells(i + 1, 10).Value = mCalcTier
        wsRes.Cells(i + 1, 11).Value = iGSRetry
        wsRes.Cells(i + 1, 12).Value = IIf(bColdMode, "cold", "warm")
        wsRes.Cells(i + 1, 13).Value = Round(ProjectElapsedHL(dSolveStart), 4)
        wsRes.Cells(i + 1, 14).Value = CStr(Now)
        wsRes.Cells(i + 1, 15).Value = Round(mCalcSecsDSCR, 4)
        wsRes.Cells(i + 1, 16).Value = Round(mCalcSecsNPP, 4)
        wsRes.Cells(i + 1, 17).Value = Round(mCalcSecsAppr, 4)
        wsRes.Cells(i + 1, 18).Value = Round(mCalcSecsFull, 4)
        wsRes.Cells(i + 1, 19).Value = iActualIters
        wsRes.Cells(i + 1, 20).Value = ClassifyConvergenceHL(dEquityPct, dIRRGap, dApprGap)
        WriteHeartbeatHL wsRes, "DONE|" & CStr(Now) & "|Project=" & arrNames(i)

    Next i

    ' --- Restore and finalize ---
    wsPI.Range(PI_PROJ_INDEX).Value = lngOriginalF2
    CalcModelCoreHL
    EnableAllSheets
    CalcOutputSheetsHL

    RestoreGoalSeekDefaultsHL
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    WriteHeartbeatHL wsRes, "COMPLETE|" & CStr(Now)
    Exit Sub

ErrHandler:
    On Error Resume Next
    wsPI.Range(PI_PROJ_INDEX).Value = lngOriginalF2
    ' Restore HoldCo so a partial-portfolio failure doesn't leave the
    ' workbook with HoldCo=0 across already-solved projects.
    wsPT.Range(PT_HOLDCO_ONOFF).Value = lngOriginalHoldCo
    CalcModelCoreHL
    EnableAllSheets
    RestoreGoalSeekDefaultsHL
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not wsRes Is Nothing Then WriteHeartbeatHL wsRes, "ERROR|" & CStr(Now)
    On Error GoTo 0
End Sub
