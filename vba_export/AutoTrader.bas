Attribute VB_Name = "AutoTrader"
Option Explicit

Private Const SHEET_DASHBOARD As String = "NewDashboard"
Private Const SHEET_CANDIDATES As String = "Candidates"
Private Const SHEET_ORDERS As String = "Orders"
Private Const CANDIDATES_REL_PATH As String = "\output\excel\candidates_nextday.csv"

Private Const DASH_HEADER_ROW As Long = 5
Private Const DASH_DATA_START As Long = 6
Private Const DASH_STATUS_CELL As String = "B2"
Private Const DASH_MAX_ORDERS_CELL As String = "B3"
Private Const DASH_SESSION_START_CELL As String = "B4"
Private Const DASH_SESSION_END_CELL As String = "B5"
Private Const DASH_SELECTED_DEFAULT_CELL As String = "B6"
Private Const DASH_REENTRY_CELL As String = "B7"
Private Const DASH_HARDSTOP_CELL As String = "B8"
Private Const DASH_LIVE_CELL As String = "B9"
Private Const DASH_ORDERMACRO_CELL As String = "B10"
Private Const DASH_CLOSE_TIME_CELL As String = "B11"
Private Const DASH_QTY_CELL As String = "B12"
Private Const DASH_TIF_CELL As String = "B13"
Private Const DASH_BUDGET_CELL As String = "B14"
Private Const DASH_LOT_STEP_CELL As String = "B15"
Private Const DASH_SLIP_BP_CELL As String = "B16"
Private Const DEFAULT_MAX_ORDERS As Long = 20
Private Const DEFAULT_SESSION_START As String = "09:00"
Private Const DEFAULT_SESSION_END As String = "09:15"
Private Const DEFAULT_SELECTED_DEFAULT As Long = 1
Private Const DEFAULT_REENTRY As Long = 0
Private Const DEFAULT_HARDSTOP As Long = 0
Private Const DEFAULT_LIVE As Long = 0
Private Const DEFAULT_CLOSE_TIME As String = "14:59:30"
Private Const DEFAULT_ORDER_QTY As Long = 100
Private Const DEFAULT_MAX_BUDGET As Double = 10000000#
Private Const DEFAULT_LOT_STEP As Long = 100
Private Const DEFAULT_SLIP_BP As Double = 30#

Private prevJ As Object
Private AutoTimer As Date
Private isRunning As Boolean
Private tradeDate As Date
Private orderCount As Long

Public Sub ButtonLoadCandidates()
    EnsureSetup
    LoadCandidates
End Sub

Public Sub ButtonPushCandidates()
    EnsureSetup
    PushCandidatesToDashboard
End Sub

Public Sub ButtonStartAuto()
    EnsureSetup
    StartAutoTrading
End Sub

Public Sub ButtonStopAuto()
    StopAutoTrading
End Sub

Public Sub ButtonRefreshNow()
    EnsureSetup
    EvaluateAndQueueOrders
End Sub

Public Sub AttachFormulasFromDashboardTemplate()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Set wsSrc = EnsureSheet("Dashboard")
    Set wsDst = EnsureSheet(SHEET_DASHBOARD)
    Dim headers As Variant
    headers = Array("現在値", "VWAP", "ATR")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        Dim h As String: h = CStr(headers(i))
        Dim cSrc As Long: cSrc = FindColumn(wsSrc, DASH_HEADER_ROW, h)
        Dim cDst As Long: cDst = FindColumn(wsDst, DASH_HEADER_ROW, h)
        If cSrc > 0 And cDst > 0 Then
            Dim f As String
            f = wsSrc.Cells(DASH_DATA_START, cSrc).Formula
            If Len(f) > 0 Then
                Dim lastRow As Long
                lastRow = wsDst.Cells(wsDst.rows.Count, 1).End(xlUp).row
                If lastRow >= DASH_DATA_START Then
                    wsDst.Range(wsDst.Cells(DASH_DATA_START, cDst), wsDst.Cells(lastRow, cDst)).Formula = f
                End If
            End If
        End If
    Next i
    wsDst.Columns.AutoFit
    MsgBox "Formulas attached from Dashboard template.", vbInformation
End Sub

Public Sub ButtonCatchUp()
    On Error Resume Next
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim root As String
    root = ThisWorkbook.path
    Dim cmd As String
    cmd = "cmd /c """ & root & "\scripts\run_nightly_now.cmd"""
    sh.Run cmd, 0, False
    MsgBox "Catch-up launched.", vbInformation
End Sub

Public Sub CancelNightlyStatusRefresh()
    On Error Resume Next
    NightlyStatus.StopNightlyMonitor
    On Error GoTo 0
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    EnsureSetup
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    ws.Range(DASH_STATUS_CELL).value = 0
End Sub

Private Sub EnsureSetup()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    If wsDash.Range(DASH_MAX_ORDERS_CELL).value = "" Then
        wsDash.Range(DASH_MAX_ORDERS_CELL).value = DEFAULT_MAX_ORDERS
    End If
    If wsDash.Range(DASH_SESSION_START_CELL).value = "" Then
        wsDash.Range(DASH_SESSION_START_CELL).value = DEFAULT_SESSION_START
    End If
    If wsDash.Range(DASH_SESSION_END_CELL).value = "" Then
        wsDash.Range(DASH_SESSION_END_CELL).value = DEFAULT_SESSION_END
    End If
    If wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value = "" Then
        wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value = DEFAULT_SELECTED_DEFAULT
    End If
    If wsDash.Range(DASH_REENTRY_CELL).value = "" Then
        wsDash.Range(DASH_REENTRY_CELL).value = DEFAULT_REENTRY
    End If
    If wsDash.Range(DASH_HARDSTOP_CELL).value = "" Then
        wsDash.Range(DASH_HARDSTOP_CELL).value = DEFAULT_HARDSTOP
    End If
    If wsDash.Range(DASH_LIVE_CELL).value = "" Then
        wsDash.Range(DASH_LIVE_CELL).value = DEFAULT_LIVE
    End If
    If wsDash.Range(DASH_CLOSE_TIME_CELL).value = "" Then
        wsDash.Range(DASH_CLOSE_TIME_CELL).value = DEFAULT_CLOSE_TIME
    End If
    If wsDash.Range(DASH_QTY_CELL).value = "" Then
        wsDash.Range(DASH_QTY_CELL).value = DEFAULT_ORDER_QTY
    End If
    If wsDash.Range(DASH_BUDGET_CELL).value = "" Then
        wsDash.Range(DASH_BUDGET_CELL).value = DEFAULT_MAX_BUDGET
    End If
    If wsDash.Range(DASH_LOT_STEP_CELL).value = "" Then
        wsDash.Range(DASH_LOT_STEP_CELL).value = DEFAULT_LOT_STEP
    End If
    If wsDash.Range(DASH_SLIP_BP_CELL).value = "" Then
        wsDash.Range(DASH_SLIP_BP_CELL).value = DEFAULT_SLIP_BP
    End If
    EnsureSheet SHEET_CANDIDATES
    Dim wsOrders As Worksheet
    Set wsOrders = EnsureSheet(SHEET_ORDERS)
    If wsOrders.Cells(1, 1).value = "" Then
        wsOrders.Range("A1:F1").value = Array("Time", "Ticker", "Side", "Price", "Qty", "Note")
    End If
    EnsureHeaders wsDash
    wsDash.Range("A2").value = "AutoTrade (0/1)"
    wsDash.Range("A3").value = "Daily Max Orders"
    wsDash.Range("A4").value = "Session Start"
    wsDash.Range("A5").value = "Session End"
    wsDash.Range("A6").value = "Selected Default (0/1)"
    wsDash.Range("A7").value = "Reentry Allowed (0/1)"
    wsDash.Range("A8").value = "Hard Stop (0/1)"
    wsDash.Range("A9").value = "Live Orders (0/1)"
    wsDash.Range("A10").value = "Order Macro Name"
    wsDash.Range("A11").value = "Close-Out Time (HH:MM:SS)"
    wsDash.Range("A12").value = "Order Quantity (Fallback)"
    wsDash.Range("A14").value = "Max Budget Per Order (JPY)"
    wsDash.Range("A15").value = "Lot Step Size (shares)"
    wsDash.Range("A16").value = "Max Slippage (bp)"
End Sub

Private Sub EnsureHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = Array("Ticker", "Selected", "SignalMode", "Session", "ATR_n", "TPk", "SLk", "J_th", "ForwardPF", "ForwardTrades", "WinCI_L", "WinCI_H", "ExpBootMean", "ExpBootLow", "ExpBootHigh", "ForwardAvgBars", "GapBucket", "GapRule", "GapSummary", "PrevClose", "PreOpenBid", "PreOpenAsk", "PreOpenMid", "LiveGapBp", "LiveGapBucket", "LiveGapAction", "DynamicQty")
    Dim baseCol As Long: baseCol = 8 ' column H
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(DASH_HEADER_ROW, baseCol + i).value = headers(i)
    Next i
End Sub

Private Sub LoadCandidates()
    Dim path As String
    path = ThisWorkbook.path & CANDIDATES_REL_PATH
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_CANDIDATES)
    ws.Cells.Clear
    If Len(Dir$(path)) = 0 Then
        MsgBox "Candidates not found: " & path, vbExclamation
        Exit Sub
    End If
    Dim f As Integer: f = FreeFile
    Dim line As String
    Dim r As Long: r = 1
    Dim values() As String
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        values = Split(line, ",")
        ws.Cells(r, 1).Resize(1, UBound(values) + 1).value = values
        r = r + 1
    Loop
    Close #f
    ws.Columns.AutoFit
    MsgBox "Loaded " & (r - 2) & " candidates.", vbInformation
End Sub

Private Sub PushCandidatesToDashboard()
    Dim wsCand As Worksheet
    Set wsCand = EnsureSheet(SHEET_CANDIDATES)
    Dim lastRow As Long
    lastRow = wsCand.Cells(wsCand.rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "Candidates sheet is empty. Load first.", vbExclamation
        Exit Sub
    End If

    Dim colTicker As Long: colTicker = FindColumn(wsCand, 1, "Ticker")
    Dim colSel As Long: colSel = FindColumn(wsCand, 1, "Selected")
    Dim colSignal As Long: colSignal = FindColumn(wsCand, 1, "SignalMode")
    Dim colSession As Long: colSession = FindColumn(wsCand, 1, "session")
    Dim colATR As Long: colATR = FindColumn(wsCand, 1, "ATR_n")
    Dim colTP As Long: colTP = FindColumn(wsCand, 1, "TPk")
    Dim colSL As Long: colSL = FindColumn(wsCand, 1, "SLk")
    Dim colJth As Long: colJth = FindColumn(wsCand, 1, "J_th")
    Dim colFpf As Long: colFpf = FindColumn(wsCand, 1, "forward_pf_eff")
    Dim colFtr As Long: colFtr = FindColumn(wsCand, 1, "forward_trades")
    Dim colWinLow As Long: colWinLow = FindColumn(wsCand, 1, "forward_win_ci_low")
    Dim colWinHigh As Long: colWinHigh = FindColumn(wsCand, 1, "forward_win_ci_high")
    Dim colExpMean As Long: colExpMean = FindColumn(wsCand, 1, "forward_exp_boot_mean")
    Dim colExpLow As Long: colExpLow = FindColumn(wsCand, 1, "forward_exp_boot_low")
    Dim colExpHigh As Long: colExpHigh = FindColumn(wsCand, 1, "forward_exp_boot_high")
    Dim colForwardAvgBars As Long: colForwardAvgBars = FindColumn(wsCand, 1, "ForwardAvgBars")
    If colForwardAvgBars = 0 Then colForwardAvgBars = FindColumn(wsCand, 1, "forward_avg_bars")
    Dim colGapBucket As Long: colGapBucket = FindColumn(wsCand, 1, "GapBucket")
    If colGapBucket = 0 Then colGapBucket = FindColumn(wsCand, 1, "forward_gap_best_bucket")
    Dim colGapRule As Long: colGapRule = FindColumn(wsCand, 1, "GapRule")
    If colGapRule = 0 Then colGapRule = FindColumn(wsCand, 1, "forward_gap_rule")
    Dim colGapSummary As Long: colGapSummary = FindColumn(wsCand, 1, "GapSummary")
    If colGapSummary = 0 Then colGapSummary = FindColumn(wsCand, 1, "forward_gap_summary")
    If colTicker = 0 Then
        MsgBox "Ticker column missing in candidates.", vbCritical
        Exit Sub
    End If

    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    Dim baseCol As Long: baseCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    If baseCol = 0 Then baseCol = 8
    Dim clearWidth As Long: clearWidth = 32
    wsDash.Range(wsDash.Cells(DASH_DATA_START, baseCol), wsDash.Cells(wsDash.rows.Count, baseCol + clearWidth)).ClearContents

    Dim colSelectedDash As Long: colSelectedDash = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim colSignalDash As Long: colSignalDash = FindColumn(wsDash, DASH_HEADER_ROW, "SignalMode")
    Dim colSessionDash As Long: colSessionDash = FindColumn(wsDash, DASH_HEADER_ROW, "Session")
    Dim colATRDash As Long: colATRDash = FindColumn(wsDash, DASH_HEADER_ROW, "ATR_n")
    Dim colTPDash As Long: colTPDash = FindColumn(wsDash, DASH_HEADER_ROW, "TPk")
    Dim colSLDash As Long: colSLDash = FindColumn(wsDash, DASH_HEADER_ROW, "SLk")
    Dim colJthDash As Long: colJthDash = FindColumn(wsDash, DASH_HEADER_ROW, "J_th")
    Dim colForwardPFDash As Long: colForwardPFDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardPF")
    Dim colForwardTradesDash As Long: colForwardTradesDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardTrades")
    Dim colWinLowDash As Long: colWinLowDash = FindColumn(wsDash, DASH_HEADER_ROW, "WinCI_L")
    Dim colWinHighDash As Long: colWinHighDash = FindColumn(wsDash, DASH_HEADER_ROW, "WinCI_H")
    Dim colExpMeanDash As Long: colExpMeanDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootMean")
    Dim colExpLowDash As Long: colExpLowDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootLow")
    Dim colExpHighDash As Long: colExpHighDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootHigh")
    Dim colForwardAvgDash As Long: colForwardAvgDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardAvgBars")
    Dim colGapBucketDash As Long: colGapBucketDash = FindColumn(wsDash, DASH_HEADER_ROW, "GapBucket")
    Dim colGapRuleDash As Long: colGapRuleDash = FindColumn(wsDash, DASH_HEADER_ROW, "GapRule")
    Dim colGapSummaryDash As Long: colGapSummaryDash = FindColumn(wsDash, DASH_HEADER_ROW, "GapSummary")
    Dim colTickerDash As Long: colTickerDash = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    Dim colDynamicQtyDash As Long: colDynamicQtyDash = FindColumn(wsDash, DASH_HEADER_ROW, "DynamicQty")

    Dim r As Long, targetRow As Long
    targetRow = DASH_DATA_START
    Dim selDef As Long: selDef = CLng(IfZero(wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value, DEFAULT_SELECTED_DEFAULT))

    For r = 2 To lastRow
        Dim ticker As String: ticker = CStr(wsCand.Cells(r, colTicker).value)
        If Len(ticker) = 0 Then GoTo NextCandidate

        wsDash.Cells(targetRow, colTickerDash).value = ticker
        Dim selVal As Variant: selVal = wsCand.Cells(r, colSel).value
        If selVal = "" Then selVal = selDef
        wsDash.Cells(targetRow, colSelectedDash).value = selVal

        If colSignal > 0 Then wsDash.Cells(targetRow, colSignalDash).value = wsCand.Cells(r, colSignal).value
        If colSession > 0 Then wsDash.Cells(targetRow, colSessionDash).value = wsCand.Cells(r, colSession).value
        If colATR > 0 Then wsDash.Cells(targetRow, colATRDash).value = wsCand.Cells(r, colATR).value
        If colTP > 0 Then wsDash.Cells(targetRow, colTPDash).value = wsCand.Cells(r, colTP).value
        If colSL > 0 Then wsDash.Cells(targetRow, colSLDash).value = wsCand.Cells(r, colSL).value
        If colJth > 0 Then wsDash.Cells(targetRow, colJthDash).value = wsCand.Cells(r, colJth).value
        If colFpf > 0 Then wsDash.Cells(targetRow, colForwardPFDash).value = wsCand.Cells(r, colFpf).value
        If colFtr > 0 Then wsDash.Cells(targetRow, colForwardTradesDash).value = wsCand.Cells(r, colFtr).value
        If colWinLow > 0 Then wsDash.Cells(targetRow, colWinLowDash).value = wsCand.Cells(r, colWinLow).value
        If colWinHigh > 0 Then wsDash.Cells(targetRow, colWinHighDash).value = wsCand.Cells(r, colWinHigh).value
        If colExpMean > 0 Then wsDash.Cells(targetRow, colExpMeanDash).value = wsCand.Cells(r, colExpMean).value
        If colExpLow > 0 Then wsDash.Cells(targetRow, colExpLowDash).value = wsCand.Cells(r, colExpLow).value
        If colExpHigh > 0 Then wsDash.Cells(targetRow, colExpHighDash).value = wsCand.Cells(r, colExpHigh).value
        If colForwardAvgBars > 0 And colForwardAvgDash > 0 Then wsDash.Cells(targetRow, colForwardAvgDash).value = wsCand.Cells(r, colForwardAvgBars).value
        If colGapBucket > 0 And colGapBucketDash > 0 Then wsDash.Cells(targetRow, colGapBucketDash).value = wsCand.Cells(r, colGapBucket).value
        If colGapRule > 0 And colGapRuleDash > 0 Then wsDash.Cells(targetRow, colGapRuleDash).value = wsCand.Cells(r, colGapRule).value
        If colGapSummary > 0 And colGapSummaryDash > 0 Then wsDash.Cells(targetRow, colGapSummaryDash).value = Left$(CStr(wsCand.Cells(r, colGapSummary).value), 255)
        If colDynamicQtyDash > 0 Then wsDash.Cells(targetRow, colDynamicQtyDash).value = ""

        targetRow = targetRow + 1
NextCandidate:
    Next r

    wsDash.Range(wsDash.Cells(DASH_HEADER_ROW, baseCol), wsDash.Cells(DASH_DATA_START - 1 + (targetRow - DASH_DATA_START), baseCol + clearWidth)).EntireColumn.AutoFit
    MsgBox "Dashboard updated with " & (targetRow - DASH_DATA_START) & " tickers.", vbInformation
End Sub

Private Sub StartAutoTrading()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    wsDash.Range(DASH_STATUS_CELL).value = 1
    If prevJ Is Nothing Then Set prevJ = CreateObject("Scripting.Dictionary")
    prevJ.RemoveAll
    tradeDate = Date
    orderCount = 0
    isRunning = True
    ScheduleNextTick 1
    MsgBox "Auto trading loop started (dry-run).", vbInformation
End Sub

Private Sub StopAutoTrading()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    wsDash.Range(DASH_STATUS_CELL).value = 0
    isRunning = False
    On Error Resume Next
    If AutoTimer <> 0 Then Application.OnTime AutoTimer, "AutoTrader.AutoTick", , False
    On Error GoTo 0
    MsgBox "Auto trading loop stopped.", vbInformation
End Sub

Private Sub ScheduleNextTick(ByVal seconds As Double)
    On Error Resume Next
    AutoTimer = Now + TimeSerial(0, 0, seconds)
    Application.OnTime AutoTimer, "AutoTrader.AutoTick"
    On Error GoTo 0
End Sub

Public Sub AutoTick()
    On Error GoTo ExitTick
    If Not isRunning Then Exit Sub
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    If wsDash.Range(DASH_STATUS_CELL).value <> 1 Then
        StopAutoTrading
        Exit Sub
    End If
    If CLng(IfZero(wsDash.Range(DASH_HARDSTOP_CELL).value, 0)) = 1 Then Exit Sub
    If Date <> tradeDate Then
        orderCount = 0
        tradeDate = Date
        If Not prevJ Is Nothing Then prevJ.RemoveAll
        If CLng(IfZero(wsDash.Range(DASH_REENTRY_CELL).value, DEFAULT_REENTRY)) = 1 Then
            ResetSelectedToDefault wsDash
        End If
    End If

    EvaluateAndQueueOrders
    ScheduleNextTick 1
ExitTick:
End Sub
Private Sub ResetSelectedToDefault(ByVal wsDash As Worksheet)
    Dim selCol As Long: selCol = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim tickerCol As Long: tickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    If selCol = 0 Or tickerCol = 0 Then Exit Sub
    Dim lastRow As Long: lastRow = wsDash.Cells(wsDash.rows.Count, tickerCol).End(xlUp).row
    Dim defVal As Long: defVal = CLng(IfZero(wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value, DEFAULT_SELECTED_DEFAULT))
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        If wsDash.Cells(r, tickerCol).value <> "" Then
            wsDash.Cells(r, selCol).value = defVal
        End If
    Next r
End Sub

Private Sub EvaluateAndQueueOrders()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    Dim startTime As Date, endTime As Date
    startTime = ParseTime(wsDash.Range(DASH_SESSION_START_CELL).value, DEFAULT_SESSION_START)
    endTime = ParseTime(wsDash.Range(DASH_SESSION_END_CELL).value, DEFAULT_SESSION_END)
    Dim tm As Date: tm = Time
    If tm < startTime Or tm > endTime Then Exit Sub

    Dim maxOrders As Long
    maxOrders = CLng(IfZero(wsDash.Range(DASH_MAX_ORDERS_CELL).value, DEFAULT_MAX_ORDERS))
    If orderCount >= maxOrders Then Exit Sub

    Dim selCol As Long: selCol = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim signalCol As Long: signalCol = FindColumn(wsDash, DASH_HEADER_ROW, "SignalMode")
    Dim sessionCol As Long: sessionCol = FindColumn(wsDash, DASH_HEADER_ROW, "Session")
    Dim tickerCol As Long: tickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    Dim priceCol As Long: priceCol = FindColumn(wsDash, DASH_HEADER_ROW, "PrevClose")
    If priceCol = 0 Then priceCol = FindColumn(wsDash, DASH_HEADER_ROW, "PreOpenMid")
    Dim jCol As Long: jCol = FindColumn(wsDash, DASH_HEADER_ROW, "J")
    Dim jthCol As Long: jthCol = FindColumn(wsDash, DASH_HEADER_ROW, "J_th")
    Dim qtyCol As Long: qtyCol = FindColumn(wsDash, DASH_HEADER_ROW, "DynamicQty")
    If selCol = 0 Or signalCol = 0 Or tickerCol = 0 Or jCol = 0 Then Exit Sub

    Dim defaultQty As Long
    defaultQty = CLng(IfZero(wsDash.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))
    Dim maxBudget As Double
    maxBudget = CDbl(IfZero(wsDash.Range(DASH_BUDGET_CELL).value, DEFAULT_MAX_BUDGET))
    Dim lotStep As Long
    lotStep = CLng(IfZero(wsDash.Range(DASH_LOT_STEP_CELL).value, DEFAULT_LOT_STEP))
    Dim slipBp As Double
    slipBp = CDbl(IfZero(wsDash.Range(DASH_SLIP_BP_CELL).value, DEFAULT_SLIP_BP))

    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.rows.Count, tickerCol).End(xlUp).row
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        Dim ticker As String: ticker = CStr(wsDash.Cells(r, tickerCol).value)
        If Len(ticker) = 0 Then GoTo UpdatePrev
        If wsDash.Cells(r, selCol).value <> 1 Then
            SetPrevJ ticker, CDbl(IfZero(wsDash.Cells(r, jCol).value, 0))
            GoTo UpdatePrev
        End If

        Dim mode As String: mode = LCase$(CStr(wsDash.Cells(r, signalCol).value))
        Dim threshold As Double: threshold = CDbl(IfZero(wsDash.Cells(r, jthCol).value, 0))
        Dim jVal As Double: jVal = CDbl(IfZero(wsDash.Cells(r, jCol).value, 0))
        Dim prevVal As Double: prevVal = GetPrevJ(ticker)

        Dim fire As Boolean
        Select Case mode
            Case "j-only"
                fire = (Abs(jVal) >= threshold)
            Case "j-cross"
                fire = (Abs(jVal) >= threshold And Abs(prevVal) < threshold)
            Case Else
                fire = False
        End Select

        If fire Then
            Dim side As String
            If jVal < 0 Then
                side = "BUY"
            Else
                side = "SELL"
            End If
            Dim px As Double
            px = CDbl(IfZero(wsDash.Cells(r, priceCol).value, 0))
            If px <= 0 Then GoTo UpdatePrev
            Dim qty As Long
            qty = ComputeDynamicQty(px, side, maxBudget, lotStep, slipBp, defaultQty)
            If qtyCol > 0 Then wsDash.Cells(r, qtyCol).value = qty
            PlaceOrder ticker, side, px, qty, mode & ":" & wsDash.Cells(r, sessionCol).value
            PlaceBracketIfAvailable wsDash, r, ticker, side, px, qty
            ScheduleCloseExit wsDash, ticker, side
            wsDash.Cells(r, selCol).value = 0
            orderCount = orderCount + 1
            If orderCount >= maxOrders Then Exit For
        End If
UpdatePrev:
        SetPrevJ ticker, CDbl(IfZero(wsDash.Cells(r, jCol).value, 0))
    Next r
End Sub

Private Sub PlaceBracketIfAvailable(ByVal ws As Worksheet, ByVal r As Long, ByVal ticker As String, ByVal side As String, ByVal entryPx As Variant, ByVal qty As Long)
    On Error Resume Next
    Dim tpCol As Long: tpCol = FindColumn(ws, DASH_HEADER_ROW, "TPk")
    Dim slCol As Long: slCol = FindColumn(ws, DASH_HEADER_ROW, "SLk")
    Dim atrPriceCol As Long: atrPriceCol = FindColumn(ws, DASH_HEADER_ROW, "ATR")
    If tpCol = 0 Or slCol = 0 Or atrPriceCol = 0 Then Exit Sub
    Dim tpK As Double: tpK = CDbl(IfZero(ws.Cells(r, tpCol).value, 0))
    Dim slK As Double: slK = CDbl(IfZero(ws.Cells(r, slCol).value, 0))
    Dim atr As Double: atr = CDbl(IfZero(ws.Cells(r, atrPriceCol).value, 0))
    If atr <= 0 Then Exit Sub
    Dim tpPx As Double, slPx As Double
    If UCase$(side) = "BUY" Then
        tpPx = CDbl(entryPx) + tpK * atr
        slPx = CDbl(entryPx) - slK * atr
    Else
        tpPx = CDbl(entryPx) - tpK * atr
        slPx = CDbl(entryPx) + slK * atr
    End If
    PlaceOrder ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), tpPx, qty, "TP"
    PlaceOrder ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), slPx, qty, "SL"
End Sub

Private Sub ScheduleCloseExit(ByVal ws As Worksheet, ByVal ticker As String, ByVal side As String)
    On Error Resume Next
    Dim t As Date
    t = ParseTime(ws.Range(DASH_CLOSE_TIME_CELL).value, DEFAULT_CLOSE_TIME)
    Application.OnTime EarliestTime:=Date + t, Procedure:="AutoTrader.CloseAtMarket", SCHEDULE:=True
    ' Store side info in Orders sheet note for later reference if必要
End Sub

Public Sub CloseAtMarket()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH_HEADER_ROW, "Ticker")
    Dim priceCol As Long: priceCol = FindColumn(ws, DASH_HEADER_ROW, "PrevClose")
    If priceCol = 0 Then priceCol = FindColumn(ws, DASH_HEADER_ROW, "PreOpenMid")
    Dim qtyCol As Long: qtyCol = FindColumn(ws, DASH_HEADER_ROW, "DynamicQty")
    If tickerCol = 0 Or priceCol = 0 Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, tickerCol).End(xlUp).row
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        Dim tkr As String: tkr = CStr(ws.Cells(r, tickerCol).value)
        If Len(tkr) > 0 Then
            Dim qty As Long
            qty = CLng(IfZero(IIf(qtyCol > 0, ws.Cells(r, qtyCol).value, 0), CLng(IfZero(ws.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))))
            PlaceOrder tkr, "FLAT", ws.Cells(r, priceCol).value, qty, "MOC"
        End If
    Next r
End Sub

Private Sub PlaceOrder(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    If qty <= 0 Then
        qty = CLng(IfZero(ws.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))
    End If
    If CLng(IfZero(ws.Range(DASH_LIVE_CELL).value, 0)) = 1 Then
        PlaceOrderLive ticker, side, price, qty, info
    Else
        PlaceOrderDryRun ticker, side, price, qty, info
    End If
End Sub

Private Sub PlaceOrderLive(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim macroName As String: macroName = CStr(IfZero(ws.Range(DASH_ORDERMACRO_CELL).value, ""))
    If Len(macroName) > 0 Then
        Application.Run macroName, ticker, side, price, qty, info
    Else
        ' fallback: dry run if macro not configured
        PlaceOrderDryRun ticker, side, price, qty, "NO_MACRO:" & info
    End If
End Sub

Private Function GetPrevJ(ByVal ticker As String) As Double
    If prevJ Is Nothing Then Set prevJ = CreateObject("Scripting.Dictionary")
    If prevJ.Exists(ticker) Then
        GetPrevJ = prevJ(ticker)
    Else
        GetPrevJ = 0
    End If
End Function

Private Sub SetPrevJ(ByVal ticker As String, ByVal value As Double)
    If prevJ Is Nothing Then Set prevJ = CreateObject("Scripting.Dictionary")
    prevJ(ticker) = value
End Sub

Private Sub PlaceOrderDryRun(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_ORDERS)
    Dim r As Long
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).row + 1
    ws.Cells(r, 1).value = Now
    ws.Cells(r, 2).value = ticker
    ws.Cells(r, 3).value = side
    ws.Cells(r, 4).value = price
    ws.Cells(r, 5).value = qty
    ws.Cells(r, 6).value = "DEMO " & info

    If qty <= 0 Then qty = CLng(IfZero(EnsureSheet(SHEET_DASHBOARD).Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))

    Dim wsPnl As Worksheet
    Set wsPnl = EnsureSheet("PnL")
    Dim pr As Long
    pr = wsPnl.Cells(wsPnl.rows.Count, 1).End(xlUp).row + 1
    If pr < 5 Then pr = 5
    wsPnl.Cells(pr, 1).value = Now
    wsPnl.Cells(pr, 3).value = "DEMO"
    wsPnl.Cells(pr, 4).value = ticker
    wsPnl.Cells(pr, 5).value = side
    wsPnl.Cells(pr, 6).value = qty
    wsPnl.Cells(pr, 7).value = price
    wsPnl.Cells(pr, 9).value = info
End Sub

Private Function ComputeDynamicQty(ByVal basePrice As Double, ByVal side As String, ByVal maxBudget As Double, ByVal lotStep As Long, ByVal slipBp As Double, ByVal fallbackQty As Long) As Long
    Dim stepSize As Long: stepSize = IIf(lotStep > 0, lotStep, 1)
    Dim budget As Double: budget = IIf(maxBudget > 0, maxBudget, 0)
    Dim qty As Long
    Dim useQty As Long: useQty = IIf(fallbackQty > 0, fallbackQty, stepSize)
    If basePrice <= 0 Or budget <= 0 Then
        ComputeDynamicQty = useQty
        Exit Function
    End If
    Dim slipFactor As Double: slipFactor = Abs(slipBp) / 10000#
    Dim worstPrice As Double
    If UCase$(side) = "BUY" Then
        worstPrice = basePrice * (1 + slipFactor)
    Else
        worstPrice = basePrice * (1 - slipFactor)
    End If
    If worstPrice <= 0 Then worstPrice = basePrice
    qty = CLng(Fix(budget / worstPrice / stepSize)) * stepSize
    If qty <= 0 Then qty = stepSize
    ComputeDynamicQty = qty
End Function

Private Function ParseTime(ByVal value As Variant, ByVal fallback As String) As Date
    On Error GoTo Fail
    If IsDate(value) Then
        ParseTime = CDate(value)
        Exit Function
    End If
    If IsDate(fallback) Then
        ParseTime = CDate(fallback)
        Exit Function
    End If
Fail:
    ParseTime = TimeValue("09:00")
End Function

Private Function FindColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal name As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim text As String
        text = CStr(ws.Cells(headerRow, c).value)
        If Len(text) > 0 Then
            If InStr(1, text, name, vbTextCompare) > 0 Then
                FindColumn = c
                Exit Function
            End If
        End If
    Next c
    FindColumn = 0
End Function

Private Function IfZero(ByVal value As Variant, ByVal fallback As Variant) As Variant
    If IsError(value) Then
        IfZero = fallback
    ElseIf IsEmpty(value) Then
        IfZero = fallback
    ElseIf value = "" Then
        IfZero = fallback
    Else
        IfZero = value
    End If
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.name = name
    End If
End Function




