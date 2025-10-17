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

Private Const DEFAULT_MAX_ORDERS As Long = 20
Private Const DEFAULT_SESSION_START As String = "09:00"
Private Const DEFAULT_SESSION_END As String = "09:15"
Private Const DEFAULT_SELECTED_DEFAULT As Long = 1
Private Const DEFAULT_REENTRY As Long = 0
Private Const DEFAULT_HARDSTOP As Long = 0
Private Const DEFAULT_LIVE As Long = 0
Private Const DEFAULT_CLOSE_TIME As String = "14:59:30"

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
                lastRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row
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
    root = ThisWorkbook.Path
    Dim cmd As String
    cmd = "cmd /c """ & root & "\scripts\run_nightly_now.cmd"""
    sh.Run cmd, 0, False
    MsgBox "Catch-up launched.", vbInformation
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    ButtonLoadCandidates
    ButtonPushCandidates
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    ws.Range(DASH_STATUS_CELL).Value = 1 ' AutoTrade ON
    ButtonStartAuto
End Sub

Private Sub EnsureSetup()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    If wsDash.Range(DASH_MAX_ORDERS_CELL).Value = "" Then
        wsDash.Range(DASH_MAX_ORDERS_CELL).Value = DEFAULT_MAX_ORDERS
    End If
    If wsDash.Range(DASH_SESSION_START_CELL).Value = "" Then
        wsDash.Range(DASH_SESSION_START_CELL).Value = DEFAULT_SESSION_START
    End If
    If wsDash.Range(DASH_SESSION_END_CELL).Value = "" Then
        wsDash.Range(DASH_SESSION_END_CELL).Value = DEFAULT_SESSION_END
    End If
    If wsDash.Range(DASH_SELECTED_DEFAULT_CELL).Value = "" Then
        wsDash.Range(DASH_SELECTED_DEFAULT_CELL).Value = DEFAULT_SELECTED_DEFAULT
    End If
    If wsDash.Range(DASH_REENTRY_CELL).Value = "" Then
        wsDash.Range(DASH_REENTRY_CELL).Value = DEFAULT_REENTRY
    End If
    If wsDash.Range(DASH_HARDSTOP_CELL).Value = "" Then
        wsDash.Range(DASH_HARDSTOP_CELL).Value = DEFAULT_HARDSTOP
    End If
    If wsDash.Range(DASH_LIVE_CELL).Value = "" Then
        wsDash.Range(DASH_LIVE_CELL).Value = DEFAULT_LIVE
    End If
    If wsDash.Range(DASH_CLOSE_TIME_CELL).Value = "" Then
        wsDash.Range(DASH_CLOSE_TIME_CELL).Value = DEFAULT_CLOSE_TIME
    End If
    EnsureSheet SHEET_CANDIDATES
    Dim wsOrders As Worksheet
    Set wsOrders = EnsureSheet(SHEET_ORDERS)
    If wsOrders.Cells(1, 1).Value = "" Then
        wsOrders.Range("A1:E1").Value = Array("Time", "Ticker", "Side", "Qty", "Note")
    End If
    EnsureHeaders wsDash
    wsDash.Range("A2").Value = "AutoTrade (0/1)"
    wsDash.Range("A3").Value = "Daily Max Orders"
    wsDash.Range("A4").Value = "Session Start"
    wsDash.Range("A5").Value = "Session End"
    wsDash.Range("A6").Value = "Selected Default (0/1)"
    wsDash.Range("A7").Value = "Reentry Allowed (0/1)"
    wsDash.Range("A8").Value = "Hard Stop (0/1)"
    wsDash.Range("A9").Value = "Live Orders (0/1)"
    wsDash.Range("A10").Value = "Order Macro Name"
    wsDash.Range("A11").Value = "Close-Out Time (HH:MM:SS)"
End Sub

Private Sub EnsureHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = Array("Selected", "SignalMode", "Session", "ATR_n", "TPk", "SLk", "J_th", "ForwardPF", "ForwardTrades", "WinCI_L", "WinCI_H", "ExpBootMean", "ExpBootLow", "ExpBootHigh")
    Dim baseCol As Long: baseCol = 24 ' column X
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(DASH_HEADER_ROW, baseCol + i).Value = headers(i)
    Next i
End Sub

Private Sub LoadCandidates()
    Dim path As String
    path = ThisWorkbook.Path & CANDIDATES_REL_PATH
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
        ws.Cells(r, 1).Resize(1, UBound(values) + 1).Value = values
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
    lastRow = wsCand.Cells(wsCand.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Candidates sheet is empty. Load first.", vbExclamation
        Exit Sub
    End If

    Dim colTicker As Long, colSignal As Long, colSession As Long
    Dim colATR As Long, colTP As Long, colSL As Long, colJth As Long, colSel As Long
    colSel = FindColumn(wsCand, 1, "Selected")
    colTicker = FindColumn(wsCand, 1, "Ticker")
    colSignal = FindColumn(wsCand, 1, "signal_mode")
    colSession = FindColumn(wsCand, 1, "session")
    colATR = FindColumn(wsCand, 1, "ATR_n")
    colTP = FindColumn(wsCand, 1, "TPk")
    colSL = FindColumn(wsCand, 1, "SLk")
    colJth = FindColumn(wsCand, 1, "J_th")
    If colTicker = 0 Then
        MsgBox "Ticker column missing in candidates.", vbCritical
        Exit Sub
    End If

    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    wsDash.Range(wsDash.Cells(DASH_DATA_START, 1), wsDash.Cells(wsDash.Rows.Count, 40)).ClearContents

    Dim r As Long, targetRow As Long
    targetRow = DASH_DATA_START
    For r = 2 To lastRow
        wsDash.Cells(targetRow, 1).Value = wsCand.Cells(r, colTicker).Value
        Dim selDef As Long: selDef = CLng(IfZero(EnsureSheet(SHEET_DASHBOARD).Range(DASH_SELECTED_DEFAULT_CELL).Value, DEFAULT_SELECTED_DEFAULT))
        Dim selVal As Variant: selVal = wsCand.Cells(r, colSel).Value
        If selVal = "" Then selVal = selDef
        wsDash.Cells(targetRow, 24).Value = selVal
        wsDash.Cells(targetRow, 25).Value = wsCand.Cells(r, colSignal).Value
        wsDash.Cells(targetRow, 26).Value = wsCand.Cells(r, colSession).Value
        wsDash.Cells(targetRow, 27).Value = wsCand.Cells(r, colATR).Value
        wsDash.Cells(targetRow, 28).Value = wsCand.Cells(r, colTP).Value
        wsDash.Cells(targetRow, 29).Value = wsCand.Cells(r, colSL).Value
        wsDash.Cells(targetRow, 30).Value = wsCand.Cells(r, colJth).Value
        wsDash.Cells(targetRow, 31).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_pf_eff")).Value
        wsDash.Cells(targetRow, 32).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_trades")).Value
        wsDash.Cells(targetRow, 33).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_win_ci_low")).Value
        wsDash.Cells(targetRow, 34).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_win_ci_high")).Value
        wsDash.Cells(targetRow, 35).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_exp_boot_mean")).Value
        wsDash.Cells(targetRow, 36).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_exp_boot_low")).Value
        wsDash.Cells(targetRow, 37).Value = wsCand.Cells(r, FindColumn(wsCand, 1, "forward_exp_boot_high")).Value
        targetRow = targetRow + 1
    Next r

    wsDash.Columns.AutoFit
    MsgBox "Dashboard updated with " & (targetRow - DASH_DATA_START) & " tickers.", vbInformation
End Sub

Private Sub StartAutoTrading()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    wsDash.Range(DASH_STATUS_CELL).Value = 1
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
    wsDash.Range(DASH_STATUS_CELL).Value = 0
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
    If wsDash.Range(DASH_STATUS_CELL).Value <> 1 Then
        StopAutoTrading
        Exit Sub
    End If
    If CLng(IfZero(wsDash.Range(DASH_HARDSTOP_CELL).Value, 0)) = 1 Then Exit Sub
    If Date <> tradeDate Then
        orderCount = 0
        tradeDate = Date
        If Not prevJ Is Nothing Then prevJ.RemoveAll
        If CLng(IfZero(wsDash.Range(DASH_REENTRY_CELL).Value, DEFAULT_REENTRY)) = 1 Then
            ResetSelectedToDefault wsDash
        End If
    End If

    EvaluateAndQueueOrders
    ScheduleNextTick 1
ExitTick:
End Sub
Private Sub ResetSelectedToDefault(ByVal wsDash As Worksheet)
    Dim selCol As Long: selCol = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    If selCol = 0 Then selCol = 24
    Dim lastRow As Long: lastRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row
    Dim defVal As Long: defVal = CLng(IfZero(wsDash.Range(DASH_SELECTED_DEFAULT_CELL).Value, DEFAULT_SELECTED_DEFAULT))
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        If wsDash.Cells(r, 1).Value <> "" Then
            wsDash.Cells(r, selCol).Value = defVal
        End If
    Next r
End Sub

Private Sub EvaluateAndQueueOrders()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    Dim startTime As Date, endTime As Date
    startTime = ParseTime(wsDash.Range(DASH_SESSION_START_CELL).Value, DEFAULT_SESSION_START)
    endTime = ParseTime(wsDash.Range(DASH_SESSION_END_CELL).Value, DEFAULT_SESSION_END)
    Dim tm As Date: tm = Time
    If tm < startTime Or tm > endTime Then Exit Sub

    Dim maxOrders As Long
    maxOrders = CLng(IfZero(wsDash.Range(DASH_MAX_ORDERS_CELL).Value, DEFAULT_MAX_ORDERS))
    If orderCount >= maxOrders Then Exit Sub

    Dim selCol As Long: selCol = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim signalCol As Long: signalCol = FindColumn(wsDash, DASH_HEADER_ROW, "SignalMode")
    Dim sessionCol As Long: sessionCol = FindColumn(wsDash, DASH_HEADER_ROW, "Session")
    Dim tickerCol As Long: tickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "コード")
    If tickerCol = 0 Then tickerCol = 1
    Dim jCol As Long: jCol = FindColumn(wsDash, DASH_HEADER_ROW, "J")
    Dim jthCol As Long: jthCol = FindColumn(wsDash, DASH_HEADER_ROW, "J_th")
    Dim priceCol As Long: priceCol = FindColumn(wsDash, DASH_HEADER_ROW, "現在値")
    If selCol = 0 Or signalCol = 0 Or tickerCol = 0 Or jCol = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, tickerCol).End(xlUp).Row
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        Dim ticker As String: ticker = CStr(wsDash.Cells(r, tickerCol).Value)
        If Len(ticker) = 0 Then GoTo UpdatePrev
        If wsDash.Cells(r, selCol).Value <> 1 Then
            SetPrevJ ticker, CDbl(IfZero(wsDash.Cells(r, jCol).Value, 0))
            GoTo UpdatePrev
        End If

        Dim mode As String: mode = LCase$(CStr(wsDash.Cells(r, signalCol).Value))
        Dim threshold As Double: threshold = CDbl(IfZero(wsDash.Cells(r, jthCol).Value, 0))
        Dim jVal As Double: jVal = CDbl(IfZero(wsDash.Cells(r, jCol).Value, 0))
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
            Dim px As Variant: px = wsDash.Cells(r, priceCol).Value
            PlaceOrder ticker, side, px, mode & ":" & wsDash.Cells(r, sessionCol).Value
            PlaceBracketIfAvailable wsDash, r, ticker, side, px
            ScheduleCloseExit wsDash, ticker, side
            wsDash.Cells(r, selCol).Value = 0
            orderCount = orderCount + 1
            If orderCount >= maxOrders Then Exit For
        End If
UpdatePrev:
        SetPrevJ ticker, CDbl(IfZero(wsDash.Cells(r, jCol).Value, 0))
    Next r
End Sub

Private Sub PlaceBracketIfAvailable(ByVal ws As Worksheet, ByVal r As Long, ByVal ticker As String, ByVal side As String, ByVal entryPx As Variant)
    On Error Resume Next
    Dim tpCol As Long: tpCol = FindColumn(ws, DASH_HEADER_ROW, "TPk")
    Dim slCol As Long: slCol = FindColumn(ws, DASH_HEADER_ROW, "SLk")
    Dim atrPriceCol As Long: atrPriceCol = FindColumn(ws, DASH_HEADER_ROW, "ATR")
    If tpCol = 0 Or slCol = 0 Or atrPriceCol = 0 Then Exit Sub
    Dim tpK As Double: tpK = CDbl(IfZero(ws.Cells(r, tpCol).Value, 0))
    Dim slK As Double: slK = CDbl(IfZero(ws.Cells(r, slCol).Value, 0))
    Dim atr As Double: atr = CDbl(IfZero(ws.Cells(r, atrPriceCol).Value, 0))
    If atr <= 0 Then Exit Sub
    Dim tpPx As Double, slPx As Double
    If UCase$(side) = "BUY" Then
        tpPx = CDbl(entryPx) + tpK * atr
        slPx = CDbl(entryPx) - slK * atr
    Else
        tpPx = CDbl(entryPx) - tpK * atr
        slPx = CDbl(entryPx) + slK * atr
    End If
    PlaceOrder ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), tpPx, "TP"
    PlaceOrder ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), slPx, "SL"
End Sub

Private Sub ScheduleCloseExit(ByVal ws As Worksheet, ByVal ticker As String, ByVal side As String)
    On Error Resume Next
    Dim t As Date
    t = ParseTime(ws.Range(DASH_CLOSE_TIME_CELL).Value, DEFAULT_CLOSE_TIME)
    Application.OnTime EarliestTime:=Date + t, Procedure:="AutoTrader.CloseAtMarket", Schedule:=True
    ' Store side info in Orders sheet note for later reference if必要
End Sub

Public Sub CloseAtMarket()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH_HEADER_ROW, "コード")
    Dim priceCol As Long: priceCol = FindColumn(ws, DASH_HEADER_ROW, "現在値")
    If tickerCol = 0 Or priceCol = 0 Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, tickerCol).End(xlUp).Row
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        Dim tkr As String: tkr = CStr(ws.Cells(r, tickerCol).Value)
        If Len(tkr) > 0 Then
            PlaceOrder tkr, "FLAT", ws.Cells(r, priceCol).Value, "MOC"
        End If
    Next r
End Sub

Private Sub PlaceOrder(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal info As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    If CLng(IfZero(ws.Range(DASH_LIVE_CELL).Value, 0)) = 1 Then
        PlaceOrderLive ticker, side, price, info
    Else
        PlaceOrderDryRun ticker, side, price, info
    End If
End Sub

Private Sub PlaceOrderLive(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal info As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim macroName As String: macroName = CStr(IfZero(ws.Range(DASH_ORDERMACRO_CELL).Value, ""))
    If Len(macroName) > 0 Then
        Application.Run macroName, ticker, side, price, info
    Else
        ' fallback: dry run if macro not configured
        PlaceOrderDryRun ticker, side, price, "NO_MACRO:" & info
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

Private Sub PlaceOrderDryRun(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal signalMode As String, ByVal sessionName As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_ORDERS)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = ticker
    ws.Cells(r, 3).Value = side
    ws.Cells(r, 4).Value = price
    ws.Cells(r, 5).Value = "DRYRUN " & signalMode & " " & sessionName
End Sub

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
        text = CStr(ws.Cells(headerRow, c).Value)
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
        EnsureSheet.Name = name
    End If
End Function
