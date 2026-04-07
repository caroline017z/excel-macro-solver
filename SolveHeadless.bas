Attribute VB_Name = "modSolveHeadless"
' ==============================================================================
'  HEADLESS CONVERGENCE SOLVER v4b — for Python COM automation
'
'  Performance optimizations:
'    1. No MsgBox / StatusBar / user dialogs
'    2. Leaves Calculation in xlCalculationManual on exit
'    3. Application.Calculate instead of 13 individual sheet calculates
'    4. Non-core sheets disabled via EnableCalculation = False
'    5. Batched GoalSeeks: NPP + Dev Fee solved before single recalc
'    6. Multi-threaded calculation enabled for parallel formula evaluation
'    7. MaxIterations=200, MAX_GS_RETRY=3 for tighter iteration control
' ==============================================================================

Option Explicit

' --- Configuration ---
Private Const MAX_ITER          As Integer = 8
Private Const MAX_GS_RETRY      As Integer = 6
Private Const EQUITY_FINAL_TOL  As Double = 0.005
Private Const IRR_TOLERANCE     As Double = 0.0003
Private Const APPR_TOLERANCE    As Double = 0.0003
Private Const DSCR_MIN          As Double = 0.5
Private Const DSCR_MAX          As Double = 5#
Private Const COL_SCAN_LIMIT    As Integer = 60

' --- Sheet names ---
Private Const SHT_PI  As String = "Project Inputs"
Private Const SHT_PT  As String = "PT Returns"

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


' ==============================================================================
'  INTERNAL HELPERS
' ==============================================================================
Private Sub CalcModelCoreHL()
    ' Full sheet-level recalc in dependency order.
    ' Individual Sheets("X").Calculate forces ALL formulas on each sheet
    ' to recalculate — essential for GoalSeek convergence with seed values.
    ' Non-core sheets are disabled via EnableCalculation=False.
    ' DoEvents after each sheet keeps COM RPC channel alive during long solves.
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
    DoEvents  ' Yield to COM message pump — prevents RPC timeout
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
    Application.MaxIterations = 1000
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
    CalcModelCoreHL
End Sub


' ==============================================================================
'  MAIN ENTRY POINT
' ==============================================================================
Public Sub SolveHeadless()

    ' --- Worksheet references ---
    Dim wsPI As Worksheet
    Dim wsPT As Worksheet
    On Error GoTo ErrHandler
    Set wsPI = ThisWorkbook.Sheets(SHT_PI)
    Set wsPT = ThisWorkbook.Sheets(SHT_PT)

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
    Dim bConverged  As Boolean
    Dim bGSok       As Boolean
    Dim dEquityPct  As Double
    Dim dPrevEqPct  As Double
    Dim dIRRGap     As Double
    Dim dApprGap    As Double
    Dim dTotalUses  As Double

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
            For iInner = 1 To MAX_GS_RETRY
                bGSok = rIRRLive.GoalSeek(Goal:=rIRRTgt.Value, ChangingCell:=rNPP)
                CalcModelCoreHL

                bGSok = rApprLive.GoalSeek(Goal:=rWACCTgt.Value, ChangingCell:=rDevFee)
                CalcModelCoreHL

                dIRRGap = Abs(rIRRLive.Value - rIRRTgt.Value)
                dApprGap = Abs(rApprLive.Value - rWACCTgt.Value)
                If dIRRGap <= IRR_TOLERANCE And dApprGap <= APPR_TOLERANCE Then Exit For
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

        Next iIter

    Next i

    ' --- Restore and finalize ---
    wsPI.Range(PI_PROJ_INDEX).Value = lngOriginalF2
    CalcModelCoreHL
    EnableAllSheets
    CalcOutputSheetsHL

    RestoreGoalSeekDefaultsHL
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    On Error Resume Next
    wsPI.Range(PI_PROJ_INDEX).Value = lngOriginalF2
    CalcModelCoreHL
    EnableAllSheets
    RestoreGoalSeekDefaultsHL
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
