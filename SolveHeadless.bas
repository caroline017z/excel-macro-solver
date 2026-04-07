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
Private Const MAX_ITER          As Integer = 8
Private Const MAX_GS_RETRY_WARM As Integer = 3
Private Const MAX_GS_RETRY_COLD As Integer = 6
Private Const GS_MAXITER_WARM   As Integer = 200
Private Const GS_MAXITER_COLD   As Integer = 1000
Private Const EQUITY_FINAL_TOL  As Double = 0.005
Private Const IRR_TOLERANCE     As Double = 0.0003
Private Const APPR_TOLERANCE    As Double = 0.0003
Private Const DSCR_MIN          As Double = 0.5
Private Const DSCR_MAX          As Double = 5#
Private Const COL_SCAN_LIMIT    As Integer = 60

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
End Sub

Private Sub WriteHeartbeatHL(ByVal wsRes As Worksheet, ByVal heartbeatText As String)
    wsRes.Range("N1").Value = heartbeatText
End Sub


' ==============================================================================
'  INTERNAL HELPERS
' ==============================================================================
Private Sub CalcModelCoreHL()
    ' Deterministic recalc ladder:
    '   Tier 1 = fastest
    '   Tier 2 = deeper dependency propagation
    '   Tier 3 = correctness guardrail for cold-start edge cases
    Select Case mCalcTier
        Case 1
            Application.Calculate
        Case 2
            Application.Calculate
            Application.Calculate
        Case Else
            Application.CalculateFull
    End Select
    DoEvents  ' Yield to COM message pump during long solve loops
End Sub

Private Sub ResetCalcTierHL()
    mCalcTier = 1
End Sub

Private Sub EscalateCalcTierHL()
    If mCalcTier < 3 Then mCalcTier = mCalcTier + 1
End Sub

Private Sub CalcOutputSheetsHL()
    Dim vSheets As Variant
    Dim vSh     As Variant
    vSheets = Array("Portfolio", "AT Returns_WIP", "Corp Model Output", _
                    "Cust Prop", "Dashboard", "Table", "Waterfall Sensitivity")
    For Each vSh In vSheets
        On Error Resume Next
        Sheets(CStr(vSh)).EnableCalculation = True
        Sheets(CStr(vSh)).Calculate
        On Error GoTo 0
    Next
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
    Application.MaxChange = 0.00001
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
    Dim bConverged  As Boolean
    Dim bGSok       As Boolean
    Dim bColdMode   As Boolean
    Dim dEquityPct  As Double
    Dim dPrevEqPct  As Double
    Dim dIRRGap     As Double
    Dim dApprGap    As Double
    Dim dTotalUses  As Double
    Dim dSolveStart As Double

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
        CalcModelCoreHL

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

        For iIter = 1 To MAX_ITER

            ' Step 1: HoldCo OFF + recalc
            rHoldCo.Value = 0
            CalcModelCoreHL

            ' Step 2: GoalSeek Min Equity = 10% (changes DSCR Multiple)
            bGSok = rEquity.GoalSeek(Goal:=rMinEqTgt.Value, ChangingCell:=rDSCR)
            If rDSCR.Value < DSCR_MIN Then rDSCR.Value = DSCR_MIN
            If rDSCR.Value > DSCR_MAX Then rDSCR.Value = DSCR_MAX
            CalcModelCoreHL

            ' Step 3: HoldCo ON + recalc
            rHoldCo.Value = 1
            CalcModelCoreHL

            ' Steps 4-5: Sequential NPP / Appraisal solve
            ' Sequential (not batched) ensures each GoalSeek sees fresh recalc
            ' values — critical for cold-start solves with seed values far from optimal.
            For iInner = 1 To iGSRetry
                bGSok = rIRRLive.GoalSeek(Goal:=rIRRTgt.Value, ChangingCell:=rNPP)
                CalcModelCoreHL

                bGSok = rApprLive.GoalSeek(Goal:=rWACCTgt.Value, ChangingCell:=rDevFee)
                CalcModelCoreHL

                dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
                dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
                If dIRRGap <= IRR_TOLERANCE And dApprGap <= APPR_TOLERANCE Then Exit For
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
                Exit For
            End If
            If Abs(dEquityPct - dPrevEqPct) < 0.000005 And iIter > 1 Then Exit For
            dPrevEqPct = dEquityPct

            ' Escalate this project from warm to cold mode only when needed.
            If Not bColdMode And iIter >= 1 Then
                bColdMode = True
                iGSRetry = MAX_GS_RETRY_COLD
                SetGoalSeekModeHL True
                EscalateCalcTierHL
            End If

        Next iIter

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
        wsRes.Cells(i + 1, 13).Value = Round(Timer - dSolveStart, 4)
        wsRes.Cells(i + 1, 14).Value = CStr(Now)
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
    CalcModelCoreHL
    EnableAllSheets
    RestoreGoalSeekDefaultsHL
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not wsRes Is Nothing Then WriteHeartbeatHL wsRes, "ERROR|" & CStr(Now)
    On Error GoTo 0
End Sub
