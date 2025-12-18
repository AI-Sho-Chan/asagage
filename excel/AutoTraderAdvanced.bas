Attribute VB_Name = "AutoTraderAdvanced"

Option Explicit



' Dashboard V2 constants (ASCII only)

Private Const DASH2_SHEET As String = "NewDashboardV2"

Private Const DASH2_HEADER_ROW As Long = 5

Private Const DASH2_DATA_START As Long = 6
Private Const DASH2_HIGHLIGHT_LAST_COL As Long = 30



Private Const DASH2_PREPLACE_FRACTION_CELL As String = "B50"

Private Const DQ As String = """"

Private Const MsoShapeRectangle As Long = 1

Private Const MsoAlignCenter As Long = -4108

Private Const DEFAULT_NKY_INITIAL_BP As Double = 10#
Private Const DEFAULT_NKY_STEADY_BP As Double = 15#
Private Const DEFAULT_ALERT_COOLDOWN_MIN As Double = 10#
Private Const SPIKE_RATIO_THRESHOLD As Double = 3#

Private gDashboardWatcher As cDashboardWatcher
Private gThresholdState As Object
Private gAlertCooldown As Object
Private gDriverConfigs As Object
Private gDriverRuntime As Object
Private gStrategyRules As Object
Private gJStats As Object
Private gBbBlockCache As Object
Private Const J_STATS_PATH As String = "state\j_stats.csv"
Private Const DRIVER_NKY As String = "NKY"
Private Const DRIVER_TOPIX As String = "TOPIX"
Private Const BB_DEFAULT_BLOCK_MINUTES As Double = 3#

' Orders sheet columns (V2)
Private Const ORD_COL_TS As Long = 1
Private Const ORD_COL_TICKER As Long = 2
Private Const ORD_COL_SIDE As Long = 3
Private Const ORD_COL_PRICE As Long = 4
Private Const ORD_COL_QTY As Long = 5
Private Const ORD_COL_MODE As Long = 6
Private Const ORD_COL_STATUS As Long = 7
Private Const ORD_COL_NOTE As Long = 8
Private Const ORD_COL_TP As Long = 9
Private Const ORD_COL_SL As Long = 10
Private Const ORD_COL_TRAIL As Long = 11
Private Const ORD_COL_FILL_TS As Long = 12
Private Const ORD_COL_FILL_PRICE As Long = 13
Private Const ORD_COL_FILL_QTY As Long = 14
Private Const ORD_COL_CLOSE_TS As Long = 15
Private Const ORD_COL_CLOSE_PRICE As Long = 16
Private Const ORD_COL_PNL_BP As Long = 17
Private Const ORD_COL_SOURCE As Long = 18

Private Enum BbRiskLevel
    bbRiskNone = 0
    bbRiskWarn = 1
    bbRiskBlock = 2
End Enum


Private Function HeaderTickerJP() As String: HeaderTickerJP = "Ticker": End Function

Private Function HeaderJValueJP() As String: HeaderJValueJP = "J": End Function

Private Function HeaderJThJP() As String: HeaderJThJP = "J_th": End Function



' Helpers

Private Function FindColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long

    Dim c As Range

    For Each c In ws.Rows(headerRow).Cells

        If Trim$(CStr(c.Value)) = headerName Then

            FindColumn = c.Column

            Exit Function

        End If

    Next c

    FindColumn = 0

End Function


Private Function FindParamColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim c As Range
    Dim aliasName As String: aliasName = GetParamHeaderAlias(headerName)
    For Each c In ws.Rows(1).Cells
        Dim cellVal As String
        cellVal = Trim$(CStr(c.Value))
        If Len(cellVal) = 0 And c.Column > 60 Then Exit For
        If StrComp(cellVal, headerName, vbTextCompare) = 0 _
            Or (aliasName <> "" And StrComp(cellVal, aliasName, vbTextCompare) = 0) Then
            FindParamColumn = c.Column
            Exit Function
        End If
    Next c
    FindParamColumn = 0
End Function

Private Function GetParamHeaderAlias(ByVal headerName As String) As String
    Select Case headerName
        Case "NKY_Code": GetParamHeaderAlias = "指標コード(日経平均)"
        Case "NKY_Last": GetParamHeaderAlias = "日経平均 現在値"
        Case "NKY_ChgPct": GetParamHeaderAlias = "日経平均 前日比率"
        Case "TOPIX_Code": GetParamHeaderAlias = "指標コード(TOPIX)"
        Case "TOPIX_Last": GetParamHeaderAlias = "TOPIX 現在値"
        Case "TOPIX_ChgPct": GetParamHeaderAlias = "TOPIX 前日比率"
        Case "Bias_bp": GetParamHeaderAlias = "バイアス閾値(bp)"
        Case "BiasSlope": GetParamHeaderAlias = "Bias補正係数"
        Case "GapSlope": GetParamHeaderAlias = "Gap補正係数"
        Case "GapBanPct": GetParamHeaderAlias = "Gap BAN 閾値(%)"
        Case "NoTradeMin": GetParamHeaderAlias = "取引停止分数"
        Case "TP_per_J": GetParamHeaderAlias = "TP/J (全体)"
        Case "SL_per_J": GetParamHeaderAlias = "SL/J (全体)"
        Case "Trail_per_J": GetParamHeaderAlias = "Trail/J (全体)"
        Case "CorrSlope": GetParamHeaderAlias = "相関補正係数"
        Case "BudgetPerTicker": GetParamHeaderAlias = "銘柄別予算(円)"
        Case "LotSize": GetParamHeaderAlias = "ロットサイズ"
        Case "NKY_TrendDay": GetParamHeaderAlias = "NKY日足トレンド"
        Case "NKY_TrendWindow": GetParamHeaderAlias = "NKY窓トレンド"
        Case "NKY_AllowedSide": GetParamHeaderAlias = "NKY許容サイド"
        Case "TOPIX_TrendDay": GetParamHeaderAlias = "TOPIX日足トレンド"
        Case "TOPIX_TrendWindow": GetParamHeaderAlias = "TOPIX窓トレンド"
        Case "TOPIX_AllowedSide": GetParamHeaderAlias = "TOPIX許容サイド"
        Case Else: GetParamHeaderAlias = ""
    End Select
End Function

Private Sub ApplyDirectionHighlight(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal entrySide As String, ByVal allowedSide As String)
    On Error Resume Next
    Dim lastCol As Long: lastCol = DASH2_HIGHLIGHT_LAST_COL
    If lastCol <= 0 Then lastCol = 30
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
    rng.Interior.ColorIndex = xlColorIndexNone
    If Len(entrySide) = 0 Then Exit Sub
    If Len(allowedSide) = 0 Or UCase$(allowedSide) = "BOTH" Then Exit Sub
    If StrComp(entrySide, allowedSide, vbTextCompare) = 0 Then
        rng.Interior.Color = RGB(235, 250, 238)
    Else
        rng.Interior.Color = RGB(255, 236, 239)
    End If
    On Error GoTo 0
End Sub

Private Sub LogVbaEvent(ByVal tag As String, ByVal message As String)
    On Error Resume Next
    Dim logPath As String
    Dim basePath As String
    ' Always prefer the workspace log dir so PowerShell can tail a stable path.
    If Len(Dir$("C:\AI\asagake", vbDirectory)) > 0 Then
        basePath = "C:\AI\asagake"
    Else
        basePath = ThisWorkbook.Path
    End If
    If Len(basePath) = 0 Then basePath = "C:\AI\asagake"
    Dim logDir As String
    logDir = basePath & "\logs"
    If Len(Dir$(logDir, vbDirectory)) = 0 Then MkDir logDir
    logPath = logDir & "\vba_events.log"
    Dim f As Integer: f = FreeFile
    Open logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & tag & "] " & message
    Close #f
End Sub

Private Function ResolveDriverCorrelation(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal driverVal As String, ByVal corrNkyCol As Long, ByVal corrTopixCol As Long) As Double
    Dim driverUpper As String
    driverUpper = UCase$(Trim$(driverVal))
    If driverUpper = DRIVER_TOPIX And corrTopixCol > 0 Then
        ResolveDriverCorrelation = ToDouble(ws.Cells(rowIndex, corrTopixCol).Value, 0#)
    Else
        ResolveDriverCorrelation = ToDouble(ws.Cells(rowIndex, corrNkyCol).Value, 0#)
    End If
End Function

Private Sub SetupTrendIndicatorCells(ByVal ws As Worksheet)
    ' Use ASCII-only labels to avoid encoding issues
    SetupSingleTrendIndicator ws, "B3:C3", "NKY Trend", "NKYTrendCell"
    SetupSingleTrendIndicator ws, "D3:E3", "TOPIX Trend", "TOPIXTrendCell"
End Sub

Private Sub SetupSingleTrendIndicator(ByVal ws As Worksheet, ByVal address As String, ByVal label As String, ByVal rangeName As String)
    Dim rng As Range
    Set rng = ws.Range(address)
    On Error Resume Next
    ws.Parent.Names(rangeName).Delete
    On Error GoTo 0
    With rng
        On Error Resume Next
        If .MergeCells Then .UnMerge
        .Merge
        On Error GoTo 0
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .Interior.Color = RGB(240, 240, 240)
        .Value = label & vbCrLf & "---"
        .Name = rangeName
    End With
End Sub

Private Sub UpdateTrendIndicators(ByVal ws As Worksheet)
    Dim nkyDayCol As Long: nkyDayCol = FindParamColumn(ws, "NKY_TrendDay")
    Dim nkyWindowCol As Long: nkyWindowCol = FindParamColumn(ws, "NKY_TrendWindow")
    Dim nkyAllowedCol As Long: nkyAllowedCol = FindParamColumn(ws, "NKY_AllowedSide")
    Dim topixDayCol As Long: topixDayCol = FindParamColumn(ws, "TOPIX_TrendDay")
    Dim topixWindowCol As Long: topixWindowCol = FindParamColumn(ws, "TOPIX_TrendWindow")
    Dim topixAllowedCol As Long: topixAllowedCol = FindParamColumn(ws, "TOPIX_AllowedSide")

    Dim nkyDay As String: If nkyDayCol > 0 Then nkyDay = NormalizeTrendState(ws.Cells(2, nkyDayCol).Value)
    Dim nkyWindow As String: If nkyWindowCol > 0 Then nkyWindow = NormalizeTrendState(ws.Cells(2, nkyWindowCol).Value)
    Dim nkyAllowed As String: If nkyAllowedCol > 0 Then nkyAllowed = UCase$(Trim$(CStr(ws.Cells(2, nkyAllowedCol).Value)))

    Dim topixDay As String: If topixDayCol > 0 Then topixDay = NormalizeTrendState(ws.Cells(2, topixDayCol).Value)
    Dim topixWindow As String: If topixWindowCol > 0 Then topixWindow = NormalizeTrendState(ws.Cells(2, topixWindowCol).Value)
    Dim topixAllowed As String: If topixAllowedCol > 0 Then topixAllowed = UCase$(Trim$(CStr(ws.Cells(2, topixAllowedCol).Value)))

    UpdateTrendIndicatorVisual ws, "NKYTrendCell", "btn_dir_nky", "NKY", nkyDay, nkyWindow, nkyAllowed
    UpdateTrendIndicatorVisual ws, "TOPIXTrendCell", "btn_dir_topix", "TOPIX", topixDay, topixWindow, topixAllowed
End Sub

Private Sub UpdateTrendIndicatorVisual(ByVal ws As Worksheet, ByVal rangeName As String, ByVal shapeName As String, ByVal label As String, ByVal dayState As String, ByVal windowState As String, ByVal allowedState As String)
    Dim displayText As String
    displayText = label & ": " & TrendStateLabel(dayState)
    If Len(windowState) > 0 Then
        displayText = displayText & " / WIN " & TrendStateLabel(windowState)
    End If
    If Len(allowedState) > 0 Then
        displayText = displayText & " / Allowed " & allowedState
    End If
    Dim fillColor As Long
    fillColor = TrendFillColor(dayState)

    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(rangeName)
    If Not rng Is Nothing Then
        rng.Value = displayText
        rng.Interior.Color = fillColor
    End If
    On Error GoTo 0
End Sub

Private Function NormalizeTrendState(ByVal raw As Variant) As String
    Dim val As String
    val = UCase$(Trim$(CStr(raw)))
    Select Case val
        Case "BUY", "SELL", "FLAT"
            NormalizeTrendState = val
        Case "UP", "UPTREND"
            NormalizeTrendState = "BUY"
        Case "DOWN", "DOWNTREND"
            NormalizeTrendState = "SELL"
        Case Else
            NormalizeTrendState = "FLAT"
    End Select
End Function

Private Function TrendStateLabel(ByVal state As String) As String
    Select Case UCase$(state)
        Case "BUY"
            TrendStateLabel = "UP"
        Case "SELL"
            TrendStateLabel = "DOWN"
        Case Else
            TrendStateLabel = "FLAT"
    End Select
End Function

Private Function TrendArrow(ByVal state As String) As String
    Select Case UCase$(state)
        Case "BUY"
            TrendArrow = "^"
        Case "SELL"
            TrendArrow = "v"
        Case Else
            TrendArrow = "-"
    End Select
End Function

Private Function TrendFillColor(ByVal state As String) As Long
    Select Case UCase$(state)
        Case "BUY"
            TrendFillColor = RGB(198, 239, 206)
        Case "SELL"
            TrendFillColor = RGB(255, 199, 206)
        Case Else
            TrendFillColor = RGB(235, 235, 235)
    End Select
End Function

Private Function GetBbBlockKey(ByVal ticker As String, ByVal session As String) As String
    GetBbBlockKey = UCase$(Trim$(ticker)) & "|" & UCase$(Trim$(session))
End Function

Private Sub EnsureBbBlockCache()
    If gBbBlockCache Is Nothing Then
        Set gBbBlockCache = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Sub ResetBbBlockCache()
    Set gBbBlockCache = Nothing
End Sub

Private Sub ActivateBbBlock(ByVal key As String)
    Dim minutes As Double
    minutes = GetStrategyRuleDouble("bb_block_minutes", BB_DEFAULT_BLOCK_MINUTES)
    If minutes <= 0# Then minutes = BB_DEFAULT_BLOCK_MINUTES
    EnsureBbBlockCache
    gBbBlockCache(key) = Now + minutes / (24# * 60#)
End Sub

Private Function IsBbBlockActive(ByVal key As String) As Boolean
    EnsureBbBlockCache
    If gBbBlockCache.Exists(key) Then
        If gBbBlockCache(key) > Now Then
            IsBbBlockActive = True
        Else
            gBbBlockCache.Remove key
        End If
    End If
End Function

Private Sub AppendSpikeEvent(ByVal ticker As String, ByVal session As String, ByVal ratioVal As Double)
    On Error GoTo ExitSub
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\analysis\j_spike_events.csv"
    Dim f As Integer
    f = FreeFile
    Dim needsHeader As Boolean
    needsHeader = (Dir$(logPath) = "")
    Open logPath For Append As #f
    If needsHeader Then Print #f, "ts,ticker,session,ratio"
    Print #f, Format$(Now, "yyyy-mm-dd HH:nn:ss") & "," & ticker & "," & session & "," & Format$(ratioVal, "0.000")
ExitSub:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Private Function GetParamDouble(ByVal ws As Worksheet, ByVal col As Long, ByVal defaultValue As Double) As Double
    On Error GoTo Fail
    If col <= 0 Then
        GetParamDouble = defaultValue
        Exit Function
    End If
    Dim val As Variant
    val = ws.Cells(2, col).Value
    If IsNumeric(val) Then
        GetParamDouble = CDbl(val)
    Else
        GetParamDouble = defaultValue
    End If
    Exit Function
Fail:
    GetParamDouble = defaultValue
End Function

Private Function ToDouble(ByVal value As Variant, ByVal defaultValue As Double) As Double
    On Error GoTo Fail
    If IsNumeric(value) Then
        ToDouble = CDbl(value)
    Else
        ToDouble = defaultValue
    End If
    Exit Function
Fail:
    ToDouble = defaultValue
End Function

Private Function RowWithFallback(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal fallbackValue As Double) As Double
    If colIndex <= 0 Then
        RowWithFallback = fallbackValue
        Exit Function
    End If
    Dim val As Variant
    val = ws.Cells(rowIndex, colIndex).Value
    If Len(Trim$(CStr(val))) = 0 Then
        RowWithFallback = fallbackValue
    ElseIf IsNumeric(val) Then
        RowWithFallback = CDbl(val)
    Else
        RowWithFallback = fallbackValue
    End If
End Function

Private Sub EnsureDriverConfigStore()
    If gDriverConfigs Is Nothing Then
        Set gDriverConfigs = CreateObject("Scripting.Dictionary")
        gDriverConfigs.CompareMode = vbTextCompare
    End If
End Sub

Private Sub EnsureDriverConfigs(ByVal ws As Worksheet)
    EnsureDriverConfigStore
    AddDriverConfig DRIVER_NKY, "NKY_Code", "NKY_Last", "NKY_TrendDay", "NKY_TrendWindow", "NKY_AllowedSide"
    AddDriverConfig DRIVER_TOPIX, "TOPIX_Code", "TOPIX_Last", "TOPIX_TrendDay", "TOPIX_TrendWindow", "TOPIX_AllowedSide"
End Sub

Private Sub AddDriverConfig(ByVal driverName As String, ByVal codeHeader As String, ByVal lastHeader As String, ByVal trendDayHeader As String, ByVal trendWindowHeader As String, ByVal allowedHeader As String)
    EnsureDriverConfigStore
    If gDriverConfigs.Exists(driverName) Then Exit Sub
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg("code_header") = codeHeader
    cfg("last_header") = lastHeader
    cfg("trend_day_header") = trendDayHeader
    cfg("trend_window_header") = trendWindowHeader
    cfg("allowed_header") = allowedHeader
    cfg("code_col") = 0
    cfg("last_col") = 0
    cfg("trend_day_col") = 0
    cfg("trend_window_col") = 0
    cfg("allowed_col") = 0
    Set gDriverConfigs(driverName) = cfg
End Sub

Private Function EnsureDriverState(ByVal driverName As String) As Object
    If gDriverRuntime Is Nothing Then
        Set gDriverRuntime = CreateObject("Scripting.Dictionary")
        gDriverRuntime.CompareMode = vbTextCompare
    End If
    If Not gDriverRuntime.Exists(driverName) Then
        Dim state As Object
        Set state = CreateObject("Scripting.Dictionary")
        state("session_date") = 0#
        state("session_open") = 0#
        Set state("history_prices") = CreateObject("Scripting.Dictionary")
        state("last_history_record") = 0#
        state("trend_day") = ""
        state("trend_window") = ""
        state("allowed_side") = "BOTH"
        state("last_allowed_side") = ""
        Set gDriverRuntime(driverName) = state
    End If
    Set EnsureDriverState = gDriverRuntime(driverName)
End Function

Private Function GetDriverColumn(ByVal ws As Worksheet, ByVal cfg As Object, ByVal cacheKey As String, ByVal headerKey As String) As Long
    Dim col As Long
    If cfg.Exists(cacheKey) Then
        col = CLng(cfg(cacheKey))
    Else
        col = 0
    End If
    If col <= 0 Then
        Dim headerName As String
        headerName = CStr(cfg(headerKey))
        col = FindParamColumn(ws, headerName)
        cfg(cacheKey) = col
    End If
    GetDriverColumn = col
End Function

Private Function NormalizeDriverName(ByVal value As String) As String
    Dim cleaned As String
    cleaned = UCase$(Trim$(value))
    If cleaned = DRIVER_TOPIX Then
        NormalizeDriverName = DRIVER_TOPIX
    Else
        NormalizeDriverName = DRIVER_NKY
    End If
End Function

Private Function GetDriverTrendDay(ByVal driverName As String) As String
    Dim state As Object
    Set state = EnsureDriverState(NormalizeDriverName(driverName))
    Dim val As String
    On Error Resume Next
    val = CStr(state("trend_day"))
    On Error GoTo 0
    If Len(val) = 0 Then val = "flat"
    GetDriverTrendDay = val
End Function

Private Function GetDriverTrendWindow(ByVal driverName As String) As String
    Dim state As Object
    Set state = EnsureDriverState(NormalizeDriverName(driverName))
    Dim val As String
    On Error Resume Next
    val = CStr(state("trend_window"))
    On Error GoTo 0
    If Len(val) = 0 Then val = "flat"
    GetDriverTrendWindow = val
End Function

Private Function GetDriverAllowedSide(ByVal driverName As String) As String
    Dim state As Object
    Set state = EnsureDriverState(NormalizeDriverName(driverName))
    Dim val As String
    On Error Resume Next
    val = CStr(state("allowed_side"))
    On Error GoTo 0
    If Len(val) = 0 Then val = "BOTH"
    GetDriverAllowedSide = val
End Function

Private Function DriverHasDownTrend(ByVal driverName As String) As Boolean
    Dim dayState As String
    Dim winState As String
    dayState = GetDriverTrendDay(driverName)
    winState = GetDriverTrendWindow(driverName)
    DriverHasDownTrend = (StrComp(dayState, "down", vbTextCompare) = 0) Or (StrComp(winState, "down", vbTextCompare) = 0)
End Function

Private Function DetermineBasePrice(ByVal vwapVal As Double, ByVal prevVal As Double) As Double
    If vwapVal > 0# Then
        DetermineBasePrice = vwapVal
    ElseIf prevVal > 0# Then
        DetermineBasePrice = prevVal
    Else
        DetermineBasePrice = 0#
    End If
End Function

Private Sub ClearRowDynamicCells(ByVal ws As Worksheet, ByVal rowIndex As Long, ParamArray cols() As Variant)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If IsNumeric(cols(i)) Then
            Dim c As Long: c = CLng(cols(i))
            If c > 0 Then
                With ws.Cells(rowIndex, c)
                    If Not .HasFormula Then
                        .ClearContents
                    End If
                End With
            End If
        End If
    Next i
End Sub

Private Function CanWriteCell(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    If colIndex <= 0 Then
        CanWriteCell = False
    Else
        CanWriteCell = Not ws.Cells(rowIndex, colIndex).HasFormula
    End If
End Function

Private Sub EnsureJStatsLoaded()
    On Error GoTo FailLoad
    If Not gJStats Is Nothing Then Exit Sub
    Dim statsPath As String
    statsPath = ThisWorkbook.Path & Application.PathSeparator & J_STATS_PATH
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(statsPath) Then Exit Sub
    Dim f As Integer
    f = FreeFile()
    Open statsPath For Input As #f
    Set gJStats = CreateObject("Scripting.Dictionary")
    Dim line As String
    Dim firstLine As Boolean: firstLine = True
    Do While Not EOF(f)
        Line Input #f, line
        If firstLine Then
            firstLine = False
        Else
            Dim parts As Variant
            parts = Split(line, ",")
            If UBound(parts) >= 4 Then
                Dim code As String: code = Trim$(UCase$(CStr(parts(0))))
                Dim sess As String: sess = Trim$(UCase$(CStr(parts(1))))
                If Len(code) > 0 And Len(sess) > 0 Then
                    Dim key As String: key = code & "|" & sess
                    Dim info(2) As Double
                    info(0) = ToDouble(parts(3), 0#)
                    info(1) = ToDouble(parts(4), 0#)
                    info(2) = ToDouble(parts(2), 0#)
                    gJStats(key) = info
                End If
            End If
        End If
    Loop
    Close #f
    Exit Sub
FailLoad:
    On Error Resume Next
    Close #f
    Set gJStats = Nothing
End Sub

Private Function GetJStatsKey(ByVal ticker As String, ByVal session As String) As String
    GetJStatsKey = UCase$(Trim$(ticker)) & "|" & UCase$(Trim$(session))
End Function

Private Function EvaluateBbRisk(ByVal ticker As String, ByVal session As String, ByVal trendWindow As String, ByVal ratioVal As Double) As BbRiskLevel
    EvaluateBbRisk = bbRiskNone
    If ratioVal <= 0# Then Exit Function
    If Len(Trim$(ticker)) = 0 Or Len(Trim$(session)) = 0 Then Exit Function
    EnsureJStatsLoaded
    If gJStats Is Nothing Then Exit Function
    Dim key As String
    key = GetJStatsKey(ticker, session)
    If Not gJStats.Exists(key) Then Exit Function
    Dim stats As Variant
    stats = gJStats(key)
    Dim samples As Double: samples = stats(2)
    Dim minSamples As Double: minSamples = GetStrategyRuleDouble("bb_min_samples", 12#)
    If samples < minSamples Then Exit Function
    Dim mu As Double: mu = stats(0)
    Dim sigma As Double: sigma = stats(1)
    If sigma <= 0# Then sigma = GetStrategyRuleDouble("bb_sigma_floor", 0.05)
    Dim trendUpper As String: trendUpper = UCase$(Trim$(trendWindow))
    Dim k As Double
    If trendUpper = "FLAT" Or Len(trendUpper) = 0 Then
        k = GetStrategyRuleDouble("bb_flat_k", 1#)
    Else
        k = GetStrategyRuleDouble("bb_trend_k", 1.3)
    End If
    Dim fatalThreshold As Double
    fatalThreshold = mu + k * sigma
    Dim warnMargin As Double
    warnMargin = GetStrategyRuleDouble("bb_warn_margin", 0.05)
    Dim blockCap As Double
    blockCap = GetStrategyRuleDouble("bb_block_ratio_cap", 0.85)
    If ratioVal < fatalThreshold Then
        If ratioVal < blockCap Then
            EvaluateBbRisk = bbRiskBlock
        Else
            EvaluateBbRisk = bbRiskWarn
        End If
    ElseIf ratioVal < fatalThreshold + warnMargin Then
        EvaluateBbRisk = bbRiskWarn
    End If
End Function

Private Function ComputeVolFactor(ByVal gapAbsPct As Double, ByVal jAbs As Double, ByVal atrVal As Double) As Double
    Dim factor As Double
    factor = 1#
    If gapAbsPct / 5# > 0.5 Then
        factor = factor + 0.5
    Else
        factor = factor + gapAbsPct / 5#
    End If
    If jAbs / 3# > 0.4 Then
        factor = factor + 0.4
    Else
        factor = factor + jAbs / 3#
    End If
    If atrVal > 0# Then
        If atrVal / 30# > 0.3 Then
            factor = factor + 0.3
        Else
            factor = factor + atrVal / 30#
        End If
    End If
    If factor < 0.7 Then factor = 0.7
    If factor > 1.8 Then factor = 1.8
    ComputeVolFactor = factor
End Function

Private Function ClassifyVolFactor(ByVal factor As Double) As String
    If factor >= 1.4 Then
        ClassifyVolFactor = "HIGH"
    ElseIf factor >= 1.0 Then
        ClassifyVolFactor = "MID"
    Else
        ClassifyVolFactor = "LOW"
    End If
End Function

Private Function ResolveRowBase(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal fallbackValue As Double) As Double
    If colIndex <= 0 Then
        ResolveRowBase = fallbackValue
    Else
        ResolveRowBase = RowWithFallback(ws, rowIndex, colIndex, fallbackValue)
    End If
End Function


Private Sub SetColumnFormula(ByVal ws As Worksheet, ByVal col As Long, ByVal fillLast As Long, ByVal formulaR1C1 As String)

    If col <= 0 Then Exit Sub

    Dim rng As Range

    Set rng = ws.Range(ws.Cells(DASH2_DATA_START, col), ws.Cells(fillLast, col))

    Dim firstCell As Range

    Set firstCell = rng.Cells(1, 1)

    On Error Resume Next

    firstCell.formulaR1C1 = formulaR1C1

    If Err.Number <> 0 Then

        Err.Clear

        firstCell.FormulaR1C1Local = formulaR1C1

    End If

    If rng.Rows.Count > 1 Then firstCell.AutoFill Destination:=rng

    On Error GoTo 0

End Sub

Private Sub EnsureDashboardWatcher()
    On Error Resume Next
    If gDashboardWatcher Is Nothing Then
        Set gDashboardWatcher = New cDashboardWatcher
        Set gDashboardWatcher.App = Application
    End If
    On Error GoTo 0
End Sub

Public Sub OnDashboardCalculate(ByVal Sh As Worksheet)
    On Error GoTo CleanExit
    If Sh Is Nothing Then Exit Sub
    If StrComp(Sh.Name, DASH2_SHEET, vbTextCompare) <> 0 Then Exit Sub
    UpdateAllDriverTrends Sh
    HandleThresholdAlerts Sh
CleanExit:
End Sub

Private Sub EnsureAlertState()
    If gThresholdState Is Nothing Then
        Set gThresholdState = CreateObject("Scripting.Dictionary")
        gThresholdState.CompareMode = vbTextCompare
    End If
    If gAlertCooldown Is Nothing Then
        Set gAlertCooldown = CreateObject("Scripting.Dictionary")
        gAlertCooldown.CompareMode = vbTextCompare
    End If
End Sub

Private Sub EnsureStrategyRules()
    If Not gStrategyRules Is Nothing Then Exit Sub
    Set gStrategyRules = CreateObject("Scripting.Dictionary")
    gStrategyRules.CompareMode = vbTextCompare
    Dim rulesPath As String
    rulesPath = ThisWorkbook.path & "\state\strategy_rules.ini"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(rulesPath) Then Exit Sub
    On Error Resume Next
    Dim stream As Object
    Set stream = fso.OpenTextFile(rulesPath, 1, False)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    Do While Not stream.AtEndOfStream
        Dim raw As String
        raw = Trim$(stream.ReadLine)
        If Len(raw) = 0 Then GoTo ContinueLoop
        If Left$(raw, 1) = "#" Then GoTo ContinueLoop
        Dim pos As Long: pos = InStr(1, raw, "=")
        If pos > 1 Then
            Dim key As String: key = Trim$(Left$(raw, pos - 1))
            Dim val As String: val = Trim$(Mid$(raw, pos + 1))
            If Len(key) > 0 Then gStrategyRules(key) = val
        End If
ContinueLoop:
    Loop
    stream.Close
End Sub

Private Function GetStrategyRule(ByVal key As String, ByVal defaultValue As String) As String
    EnsureStrategyRules
    If gStrategyRules Is Nothing Then
        GetStrategyRule = defaultValue
    ElseIf gStrategyRules.Exists(key) Then
        GetStrategyRule = CStr(gStrategyRules(key))
    Else
        GetStrategyRule = defaultValue
    End If
End Function

Private Function GetStrategyRuleDouble(ByVal key As String, ByVal defaultValue As Double) As Double
    Dim raw As String
    raw = GetStrategyRule(key, CStr(defaultValue))
    On Error GoTo Fail
    GetStrategyRuleDouble = CDbl(raw)
    Exit Function
Fail:
    GetStrategyRuleDouble = defaultValue
End Function

Private Function CanFireAlert(ByVal key As String) As Boolean
    EnsureAlertState
    Dim cooldownMinutes As Double
    cooldownMinutes = GetStrategyRuleDouble("alert_cooldown_min", DEFAULT_ALERT_COOLDOWN_MIN)
    Dim cooldownDays As Double
    cooldownDays = (cooldownMinutes / 60#) / 24#
    If gAlertCooldown.Exists(key) Then
        Dim lastVal As Double
        lastVal = CDbl(gAlertCooldown(key))
        If Now - lastVal < cooldownDays Then
            CanFireAlert = False
            Exit Function
        End If
    End If
    gAlertCooldown(key) = CDbl(Now)
    CanFireAlert = True
End Function

Private Sub RaiseThresholdAlert(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal ticker As String, ByVal side As String, ByVal ratio As Double)
    Dim message As String
    message = "THRESHOLD " & ticker & " " & side & " ratio " & Format$(ratio, "0.00")
    Application.StatusBar = message
    Dim entryStatusCol As Long
    entryStatusCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryStatus")
    If entryStatusCol > 0 Then
        ws.Cells(rowIndex, entryStatusCol).Value = "BLOCKED_ALERT " & Format$(Now, "HH:MM:SS")
    End If
End Sub

Private Sub HandleThresholdAlerts(ByVal ws As Worksheet)
    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderTickerJP())
    Dim sideCol As Long: sideCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySide")
    Dim ratioCol As Long: ratioCol = FindColumn(ws, DASH2_HEADER_ROW, "J_ratio")
    Dim jCol As Long: jCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJValueJP())
    Dim jthCol As Long: jthCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJThJP())
    If ratioCol = 0 Or tickerCol = 0 Or jCol = 0 Or jthCol = 0 Then Exit Sub
    EnsureAlertState
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = DASH2_DATA_START To lastRow
        Dim ticker As String
        ticker = Trim$(CStr(ws.Cells(r, tickerCol).Value))
        If Len(ticker) = 0 Then Exit For
        Dim entrySide As String
        entrySide = Trim$(CStr(ws.Cells(r, sideCol).Value))
        If Len(entrySide) = 0 Then
            Dim jVal As Double: jVal = ToDouble(ws.Cells(r, jCol).Value, 0#)
            If jVal < 0# Then
                entrySide = "BUY"
            ElseIf jVal > 0# Then
                entrySide = "SELL"
            Else
                entrySide = ""
            End If
        End If
        If Len(entrySide) = 0 Then GoTo ContinueLoop
        Dim ratio As Double
        ratio = ToDouble(ws.Cells(r, ratioCol).Value, 0#)
        Dim key As String
        key = ticker & "|" & entrySide
        Dim prevAbove As Boolean
        prevAbove = False
        If gThresholdState.Exists(key) Then
            prevAbove = CBool(gThresholdState(key))
        End If
        Dim nowAbove As Boolean
        nowAbove = (ratio >= 1#)
        If nowAbove Then
            If Not prevAbove Then
                If CanFireAlert(key) Then
                    RaiseThresholdAlert ws, r, ticker, entrySide, ratio
                End If
            End If
            gThresholdState(key) = True
        Else
            gThresholdState(key) = False
        End If
ContinueLoop:
    Next r
End Sub

Private Sub UpdateAllDriverTrends(ByVal ws As Worksheet)
    EnsureDriverConfigs ws
    If gDriverConfigs Is Nothing Then Exit Sub
    Dim driverName As Variant
    For Each driverName In gDriverConfigs.Keys
        UpdateDriverTrend ws, CStr(driverName)
    Next driverName
End Sub

Private Sub UpdateDriverTrend(ByVal ws As Worksheet, ByVal driverName As String)
    EnsureDriverConfigs ws
    If gDriverConfigs Is Nothing Then Exit Sub
    If Not gDriverConfigs.Exists(driverName) Then Exit Sub
    Dim cfg As Object
    Set cfg = gDriverConfigs(driverName)
    Dim lastCol As Long
    lastCol = GetDriverColumn(ws, cfg, "last_col", "last_header")
    If lastCol <= 0 Then Exit Sub
    Dim currentPrice As Double
    currentPrice = ToDouble(ws.Cells(2, lastCol).Value, 0#)
    If currentPrice <= 0# Then Exit Sub

    Dim state As Object
    Set state = EnsureDriverState(driverName)
    Dim sessionDate As Variant
    Dim sessionOpen As Double
    sessionOpen = 0#
    On Error Resume Next
    sessionDate = state("session_date")
    sessionOpen = CDbl(state("session_open"))
    On Error GoTo 0
    If IsEmpty(sessionDate) Then sessionDate = 0#
    If sessionDate <> Date Or sessionOpen <= 0# Then
        state("session_date") = Date
        state("session_open") = currentPrice
        On Error Resume Next
        Set state("history_prices") = Nothing
        On Error GoTo 0
        state("last_history_record") = 0#
        sessionOpen = currentPrice
    End If

    Dim historyPrices As Collection
    On Error Resume Next
    Set historyPrices = state("history_prices")
    On Error GoTo 0
    If historyPrices Is Nothing Then
        Set historyPrices = New Collection
    End If

    Dim lastRec As Date
    On Error Resume Next
    lastRec = state("last_history_record")
    On Error GoTo 0
    Dim shouldAppend As Boolean
    If lastRec = 0# Then
        shouldAppend = True
    ElseIf Now - lastRec >= TimeSerial(0, 1, 0) Then
        shouldAppend = True
    End If
    If shouldAppend Then
        historyPrices.Add currentPrice
        state("last_history_record") = Now
    End If
    Do While historyPrices.Count > 15
        historyPrices.Remove 1
    Loop
    Set state("history_prices") = historyPrices

    Dim openRet As Double
    If sessionOpen > 0# Then
        openRet = (currentPrice - sessionOpen) / sessionOpen * 10000#
    Else
        openRet = 0#
    End If

    Dim historyCount As Long
    historyCount = historyPrices.Count

    Dim earliestPrice As Double
    If historyCount >= 1 Then
        earliestPrice = ToDouble(historyPrices(1), currentPrice)
    Else
        earliestPrice = sessionOpen
    End If

    Dim windowRet As Double
    If earliestPrice > 0# Then
        windowRet = (currentPrice - earliestPrice) / earliestPrice * 10000#
    Else
        windowRet = openRet
    End If

    Dim initialThreshold As Double
    initialThreshold = GetStrategyRuleDouble("nky_initial_bp", DEFAULT_NKY_INITIAL_BP)
    Dim steadyThreshold As Double
    steadyThreshold = GetStrategyRuleDouble("nky_steady_bp", DEFAULT_NKY_STEADY_BP)
    Dim threshold As Double
    If historyCount < 15 Then
        threshold = initialThreshold
    Else
        threshold = steadyThreshold
    End If

    Dim trendDay As String
    If Abs(openRet) >= threshold Then
        trendDay = IIf(openRet > 0#, "up", "down")
    Else
        trendDay = "flat"
    End If

    Dim trendWindow As String
    If historyCount >= 2 And Abs(windowRet) >= threshold Then
        trendWindow = IIf(windowRet > 0#, "up", "down")
    Else
        trendWindow = trendDay
    End If

    Dim allowedSide As String
    If trendWindow = "up" Then
        allowedSide = "BUY"
    ElseIf trendWindow = "down" Then
        allowedSide = "SELL"
    ElseIf trendDay = "up" Then
        allowedSide = "BUY"
    ElseIf trendDay = "down" Then
        allowedSide = "SELL"
    Else
        allowedSide = "BOTH"
    End If

    Dim dayCol As Long: dayCol = GetDriverColumn(ws, cfg, "trend_day_col", "trend_day_header")
    Dim windowCol As Long: windowCol = GetDriverColumn(ws, cfg, "trend_window_col", "trend_window_header")
    Dim sideCol As Long: sideCol = GetDriverColumn(ws, cfg, "allowed_col", "allowed_header")
    If dayCol > 0 Then ws.Cells(2, dayCol).Value = trendDay
    If windowCol > 0 Then ws.Cells(2, windowCol).Value = trendWindow
    If sideCol > 0 Then ws.Cells(2, sideCol).Value = allowedSide

    Dim prevAllowed As String
    On Error Resume Next
    prevAllowed = CStr(state("last_allowed_side"))
    On Error GoTo 0
    If StrComp(prevAllowed, allowedSide, vbTextCompare) <> 0 Then
        state("last_allowed_side") = allowedSide
        If Not IsDemoMode() Then CancelOppositeOrders allowedSide
    End If

    state("trend_day") = trendDay
    state("trend_window") = trendWindow
    state("allowed_side") = allowedSide
End Sub

Private Sub CancelOppositeOrders(ByVal allowedSide As String)
    If Len(allowedSide) = 0 Or allowedSide = "BOTH" Then Exit Sub
    Dim cancelSide As String
    If StrComp(allowedSide, "BUY", vbTextCompare) = 0 Then
        cancelSide = "SELL"
    ElseIf StrComp(allowedSide, "SELL", vbTextCompare) = 0 Then
        cancelSide = "BUY"
    Else
        Exit Sub
    End If
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("Orders")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim modeVal As String
        modeVal = LCase$(Trim$(CStr(sh.Cells(r, 6).Value)))
        Dim statusVal As String
        statusVal = UCase$(Trim$(CStr(sh.Cells(r, 7).Value)))
        If modeVal <> "preplace" And modeVal <> "preplace_demo" Then
            GoTo ContinueCancelLoop
        End If
        If statusVal <> "PENDING" And statusVal <> "ORDERED" Then
            GoTo ContinueCancelLoop
        End If
        Dim orderSide As String
        orderSide = UCase$(CStr(sh.Cells(r, 3).Value))
        If orderSide = UCase$(cancelSide) Then
            sh.Cells(r, 7).Value = "CANCELLED_AUTO"
        End If
ContinueCancelLoop:
    Next r
End Sub



Private Function BuildR1C1Ref(ByVal sourceCol As Long, ByVal targetCol As Long) As String
    Dim d As Long: d = sourceCol - targetCol
    If d = 0 Then
        BuildR1C1Ref = "RC"
    Else
        BuildR1C1Ref = "RC[" & CStr(d) & "]"
    End If
End Function

Private Function ParseCsvLine(ByVal line As String) As Variant

    Dim values As Collection

    Set values = New Collection

    Dim current As String

    Dim i As Long

    Dim ch As String

    Dim inQuotes As Boolean

    For i = 1 To Len(line)

        ch = Mid$(line, i, 1)

        If ch = """" Then

            If inQuotes And i < Len(line) And Mid$(line, i + 1, 1) = """" Then

                current = current & """"

                i = i + 1

            Else

                inQuotes = Not inQuotes

            End If

        ElseIf ch = "," And Not inQuotes Then

            values.Add current

            current = vbNullString

        Else

            current = current & ch

        End If

    Next i

    values.Add current

    Dim arr() As String

    ReDim arr(0 To values.Count - 1)

    For i = 1 To values.Count

        arr(i - 1) = values(i)

    Next i

    ParseCsvLine = arr

End Function



' ----------------------------------------------------------------------------

' Setup/Install

' ----------------------------------------------------------------------------

Public Sub SetupNewDashboardV2()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.name = DASH2_SHEET

    End If



    Dim r As Long: r = DASH2_HEADER_ROW

    ws.Cells(r, 1).Value = HeaderTickerJP()

    ws.Cells(r, 2).Value = "Selected"

    ws.Cells(r, 3).Value = HeaderJValueJP()

    ws.Cells(r, 4).Value = HeaderJThJP()

    ws.Cells(r, 5).Value = "EntryBuyPx"

    ws.Cells(r, 6).Value = "EntrySellPx"

    ws.Cells(r, 7).Value = "EntrySide"

    ws.Cells(r, 8).Value = "EntryStatus"

    ws.Cells(r, 9).Value = "TP_price"

    ws.Cells(r, 10).Value = "SL_price"

    ws.Cells(r, 11).Value = "StopTrail"

    ws.Cells(r, 12).Value = "BestBid"

    ws.Cells(r, 13).Value = "BestAsk"

    ws.Cells(r, 14).Value = "PrevClose"

    ws.Cells(r, 15).Value = "VWAP"

    ws.Cells(r, 16).Value = "Gap_bp"

    ws.Cells(r, 17).Value = "CorrNKY"

    ws.Cells(r, 18).Value = "OrderQtyPlan"

    ws.Cells(r, 19).Value = "BiasSlope_row"

    ws.Cells(r, 20).Value = "GapSlope_row"

    ws.Cells(r, 21).Value = "CorrSlope_row"

    ws.Cells(r, 22).Value = "TP_per_J_row"

    ws.Cells(r, 23).Value = "SL_per_J_row"

    ws.Cells(r, 24).Value = "Trail_per_J_row"

    ws.Cells(r, 25).Value = "TP_per_J_eff"

    ws.Cells(r, 26).Value = "SL_per_J_eff"

    ws.Cells(r, 27).Value = "Trail_per_J_eff"

    ws.Cells(r, 28).Value = "VolatilityTag"

End Sub



Public Sub InstallRealtimeFormulasV2()

    ' Intentionally minimal (no Rss formulas here; keep offline rule)

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    EnsureDashboardWatcher
    UpdateAllDriverTrends ws

    Dim gapCol As Long: gapCol = FindColumn(ws, DASH2_HEADER_ROW, "Gap_bp")

    Dim prevCol As Long: prevCol = FindColumn(ws, DASH2_HEADER_ROW, "PrevClose")

    Dim vwapCol As Long: vwapCol = FindColumn(ws, DASH2_HEADER_ROW, "VWAP")

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If gapCol > 0 And prevCol > 0 And vwapCol > 0 Then

        Dim f As String

        ' Single-line (use DQ constant for literal quotes)
        f = "=IF(OR(RC[" & CStr(vwapCol - gapCol) & "]=" & DQ & DQ & ",RC[" & CStr(prevCol - gapCol) & "]=" & DQ & DQ & ")," & DQ & DQ & ",(RC[" & CStr(vwapCol - gapCol) & "]-RC[" & CStr(prevCol - gapCol) & "])/RC[" & CStr(prevCol - gapCol) & "]*10000)"

        SetColumnFormula ws, gapCol, lastRow, f

    End If

    Dim ratioCol As Long: ratioCol = FindColumn(ws, DASH2_HEADER_ROW, "J_ratio")
    Dim jCol As Long: jCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJValueJP())
    Dim jthCol As Long: jthCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJThJP())
    If ratioCol > 0 And jCol > 0 And jthCol > 0 Then
        Dim ratioFormula As String
        ratioFormula = "=IF(OR(RC[" & (jCol - ratioCol) & "]="""",RC[" & (jthCol - ratioCol) & "]="""",N(RC[" & (jthCol - ratioCol) & "])=0),"""",ABS(RC[" & (jCol - ratioCol) & "])/ABS(RC[" & (jthCol - ratioCol) & "]))"
        SetColumnFormula ws, ratioCol, lastRow, ratioFormula
        ApplyThresholdFormatting ws, jCol, ratioCol, lastRow
    End If

    Dim nkyDayParamCol As Long: nkyDayParamCol = FindParamColumn(ws, "NKY_TrendDay")
    Dim nkyWindowParamCol As Long: nkyWindowParamCol = FindParamColumn(ws, "NKY_TrendWindow")
    Dim nkyAllowedParamCol As Long: nkyAllowedParamCol = FindParamColumn(ws, "NKY_AllowedSide")

    Dim nkyDayCol As Long: nkyDayCol = FindColumn(ws, DASH2_HEADER_ROW, "NKY_day_trend")
    If nkyDayCol > 0 And nkyDayParamCol > 0 Then
        Dim dayFormula As String
        dayFormula = "=IF(R2C" & nkyDayParamCol & "="""","",R2C" & nkyDayParamCol & ")"
        SetColumnFormula ws, nkyDayCol, lastRow, dayFormula
    End If

    Dim nkyWindowCol As Long: nkyWindowCol = FindColumn(ws, DASH2_HEADER_ROW, "NKY_window_trend")
    If nkyWindowCol > 0 And nkyWindowParamCol > 0 Then
        Dim windowFormula As String
        windowFormula = "=IF(R2C" & nkyWindowParamCol & "="""","",R2C" & nkyWindowParamCol & ")"
        SetColumnFormula ws, nkyWindowCol, lastRow, windowFormula
    End If

    Dim nkyAllowedCol As Long: nkyAllowedCol = FindColumn(ws, DASH2_HEADER_ROW, "NKY_allowed_side")
    If nkyAllowedCol > 0 And nkyAllowedParamCol > 0 Then
        Dim allowedFormula As String
        allowedFormula = "=IF(R2C" & nkyAllowedParamCol & "="""","",R2C" & nkyAllowedParamCol & ")"
        SetColumnFormula ws, nkyAllowedCol, lastRow, allowedFormula
    End If

End Sub

Private Sub ApplyThresholdFormatting(ByVal ws As Worksheet, ByVal jCol As Long, ByVal ratioCol As Long, ByVal lastRow As Long)
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(DASH2_DATA_START, jCol), ws.Cells(lastRow, jCol))
    rng.FormatConditions.Delete
    Dim offset As Long: offset = ratioCol - jCol
    Dim fcSoft As FormatCondition
    Set fcSoft = rng.FormatConditions.Add(Type:=2, Formula1:="=AND(RC<>"""",RC[" & offset & "]<>"""",N(RC[" & offset & "])>=0.8,N(RC[" & offset & "])<1)")
    fcSoft.Interior.Color = RGB(226, 239, 218)
    Dim fcHard As FormatCondition
    Set fcHard = rng.FormatConditions.Add(Type:=2, Formula1:="=AND(RC<>"""",RC[" & offset & "]<>"""",N(RC[" & offset & "])>=1)")
    fcHard.Interior.Color = RGB(182, 215, 168)
    On Error GoTo 0
End Sub



' ----------------------------------------------------------------------------

' Signals and Orders

' ----------------------------------------------------------------------------
Private Function LoadSlippageOverrides() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fullPath As String
    fullPath = ThisWorkbook.Path & "\output\excel\slippage_overrides.csv"
    If Not fso.FileExists(fullPath) Then
        Set LoadSlippageOverrides = dict
        Exit Function
    End If

    Const ForReading = 1
    Const TristateTrue = -1

    Dim stream As Object
    On Error Resume Next
    Set stream = fso.OpenTextFile(fullPath, ForReading, False, TristateTrue)
    If Err.Number <> 0 Then
        Err.Clear
        Set LoadSlippageOverrides = dict
        Exit Function
    End If
    On Error GoTo 0

    If Not stream.AtEndOfStream Then
        stream.ReadLine ' skip header
    End If

    Do While Not stream.AtEndOfStream
        Dim line As String
        line = Trim$(stream.ReadLine)
        If Len(line) = 0 Then GoTo ContinueLoop

        Dim parts() As String
        parts = Split(line, ",")
        If UBound(parts) < 3 Then GoTo ContinueLoop

        Dim plan As String: plan = Trim$(parts(0))
        Dim sessionVal As String: sessionVal = Trim$(parts(1))
        Dim bufStr As String: bufStr = Trim$(parts(2))
        Dim bufferBp As Double
        On Error Resume Next
        bufferBp = CDbl(bufStr)
        If Err.Number <> 0 Then
            Err.Clear
            bufferBp = 0
        End If
        On Error GoTo 0

        Dim modeVal As String: modeVal = ""
        Dim planParts() As String
        planParts = Split(plan, "_")
        If UBound(planParts) >= 0 Then
            modeVal = planParts(UBound(planParts))
        End If

        If Len(sessionVal) > 0 And Len(modeVal) > 0 Then
            Dim key As String
            key = sessionVal & "|" & modeVal
            If Not dict.Exists(key) Then
                dict.Add key, bufferBp
            End If
        End If
ContinueLoop:
    Loop
    stream.Close

    Set LoadSlippageOverrides = dict
End Function

Private Function BuildAdjustedPriceFormula(ByVal baseRef As String, ByVal bufferFrac As Double, ByVal isBuy As Boolean) As String
    If Len(baseRef) = 0 Then
        BuildAdjustedPriceFormula = "=" & DQ & DQ
        Exit Function
    End If

    If bufferFrac <= 0 Then
        BuildAdjustedPriceFormula = "=" & baseRef
    Else
        Dim formatted As String
        formatted = Replace(Format$(bufferFrac, "0.################"), ",", ".")
        Dim factorExpr As String
        If isBuy Then
            factorExpr = "(1-" & formatted & ")"
        Else
            factorExpr = "(1+" & formatted & ")"
        End If
        BuildAdjustedPriceFormula = "=IF(" & baseRef & "=" & DQ & DQ & "," & DQ & DQ & "," & baseRef & "*" & factorExpr & ")"
    End If
End Function

Private Function BuildNoteFormula(ByVal sessionRef As String, ByVal modeRef As String, ByVal noteAppend As String) As String
    Dim baseNote As String
    If sessionRef <> "" Or modeRef <> "" Then
        baseNote = "=" & DQ & "session=" & DQ
        If sessionRef <> "" Then
            baseNote = baseNote & " & IF(ISBLANK(" & sessionRef & ")," & DQ & DQ & "," & sessionRef & ")"
        Else
            baseNote = baseNote & " & " & DQ & DQ
        End If
        baseNote = baseNote & " & " & DQ & ";mode=" & DQ
        If modeRef <> "" Then
            baseNote = baseNote & " & IF(ISBLANK(" & modeRef & ")," & DQ & DQ & "," & modeRef & ")"
        Else
            baseNote = baseNote & " & " & DQ & DQ
        End If
    Else
        baseNote = ""
    End If

    If Len(noteAppend) > 0 Then
        If baseNote = "" Then
            BuildNoteFormula = "=" & DQ & noteAppend & DQ
        Else
            BuildNoteFormula = baseNote & " & " & DQ & ";" & DQ & " & " & DQ & noteAppend & DQ
        End If
    Else
        BuildNoteFormula = baseNote
    End If
End Function

Private Function GetSlippageKey(ByVal sessionVal As String, ByVal modeVal As String) As String
    If Len(sessionVal) = 0 Or Len(modeVal) = 0 Then
        GetSlippageKey = ""
    Else
        GetSlippageKey = sessionVal & "|" & modeVal
    End If
End Function

Public Sub ApplyDynamicSignalsV2()

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    EnsureParamFormulas ws
    EnsureDashboardWatcher
    UpdateAllDriverTrends ws
    Set gJStats = Nothing
    EnsureJStatsLoaded

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DASH2_DATA_START Then Exit Sub

    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderTickerJP())
    Dim jCol As Long: jCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJValueJP())
    Dim jthCol As Long: jthCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJThJP())
    Dim jthBaseCol As Long: jthBaseCol = FindColumn(ws, DASH2_HEADER_ROW, "J_th_base")
    Dim vwapCol As Long: vwapCol = FindColumn(ws, DASH2_HEADER_ROW, "VWAP")
    Dim prevCol As Long: prevCol = FindColumn(ws, DASH2_HEADER_ROW, "PrevClose")
    Dim lastCol As Long: lastCol = FindColumn(ws, DASH2_HEADER_ROW, "Last")
    Dim gapCol As Long: gapCol = FindColumn(ws, DASH2_HEADER_ROW, "Gap_bp")
    Dim corrCol As Long: corrCol = FindColumn(ws, DASH2_HEADER_ROW, "CorrNKY")
    Dim corrTopixCol As Long: corrTopixCol = FindColumn(ws, DASH2_HEADER_ROW, "CorrTOPIX")
    Dim atrCol As Long: atrCol = FindColumn(ws, DASH2_HEADER_ROW, "ATR_n")
    Dim eBuyCol As Long: eBuyCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryBuyPx")
    Dim eSellCol As Long: eSellCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySellPx")
    Dim sideCol As Long: sideCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySide")
    Dim qtyCol As Long: qtyCol = FindColumn(ws, DASH2_HEADER_ROW, "OrderQtyPlan")
    Dim tpCol As Long: tpCol = FindColumn(ws, DASH2_HEADER_ROW, "TP_price")
    Dim slCol As Long: slCol = FindColumn(ws, DASH2_HEADER_ROW, "SL_price")
    Dim trailCol As Long: trailCol = FindColumn(ws, DASH2_HEADER_ROW, "StopTrail")
    Dim biasSlopeRowCol As Long: biasSlopeRowCol = FindColumn(ws, DASH2_HEADER_ROW, "BiasSlope_row")
    Dim gapSlopeRowCol As Long: gapSlopeRowCol = FindColumn(ws, DASH2_HEADER_ROW, "GapSlope_row")
    Dim corrSlopeRowCol As Long: corrSlopeRowCol = FindColumn(ws, DASH2_HEADER_ROW, "CorrSlope_row")
    Dim tpRowCol As Long: tpRowCol = FindColumn(ws, DASH2_HEADER_ROW, "TP_per_J_row")
    Dim slRowCol As Long: slRowCol = FindColumn(ws, DASH2_HEADER_ROW, "SL_per_J_row")
    Dim trailRowCol As Long: trailRowCol = FindColumn(ws, DASH2_HEADER_ROW, "Trail_per_J_row")
    Dim budgetFactorCol As Long: budgetFactorCol = FindColumn(ws, DASH2_HEADER_ROW, "BudgetFactor_row")
    Dim tpEffCol As Long: tpEffCol = FindColumn(ws, DASH2_HEADER_ROW, "TP_per_J_eff")
    Dim slEffCol As Long: slEffCol = FindColumn(ws, DASH2_HEADER_ROW, "SL_per_J_eff")
    Dim trailEffCol As Long: trailEffCol = FindColumn(ws, DASH2_HEADER_ROW, "Trail_per_J_eff")
    Dim entryStatusCol As Long: entryStatusCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryStatus")
    Dim selCol As Long: selCol = FindColumn(ws, DASH2_HEADER_ROW, "Selected")
    Dim batchKindCol As Long: batchKindCol = FindColumn(ws, DASH2_HEADER_ROW, "BatchKind")
    Dim volTagCol As Long: volTagCol = FindColumn(ws, DASH2_HEADER_ROW, "VolatilityTag")
    Dim trendWindowCol As Long: trendWindowCol = FindColumn(ws, DASH2_HEADER_ROW, "trend_window")
    Dim trendDriverCol As Long: trendDriverCol = FindColumn(ws, DASH2_HEADER_ROW, "trend_driver")
    Dim driverDayCol As Long: driverDayCol = FindColumn(ws, DASH2_HEADER_ROW, "driver_day_trend")
    Dim driverWindowCol2 As Long: driverWindowCol2 = FindColumn(ws, DASH2_HEADER_ROW, "driver_window_trend")
    Dim driverAllowedCol As Long: driverAllowedCol = FindColumn(ws, DASH2_HEADER_ROW, "driver_allowed_side")
    Dim sessionCol As Long: sessionCol = FindColumn(ws, DASH2_HEADER_ROW, "session")
    Dim modeCol As Long: modeCol = FindColumn(ws, DASH2_HEADER_ROW, "SignalMode")

    If tickerCol = 0 Or jCol = 0 Or jthCol = 0 Or jthBaseCol = 0 Then Exit Sub

    Dim biasParamCol As Long: biasParamCol = FindParamColumn(ws, "Bias_bp")
    Dim biasSlopeParamCol As Long: biasSlopeParamCol = FindParamColumn(ws, "BiasSlope")
    Dim gapSlopeParamCol As Long: gapSlopeParamCol = FindParamColumn(ws, "GapSlope")
    Dim corrSlopeParamCol As Long: corrSlopeParamCol = FindParamColumn(ws, "CorrSlope")
    Dim gapBanParamCol As Long: gapBanParamCol = FindParamColumn(ws, "GapBanPct")
    Dim tpParamCol As Long: tpParamCol = FindParamColumn(ws, "TP_per_J")
    Dim slParamCol As Long: slParamCol = FindParamColumn(ws, "SL_per_J")
    Dim trailParamCol As Long: trailParamCol = FindParamColumn(ws, "Trail_per_J")
    Dim budgetParamCol As Long: budgetParamCol = FindParamColumn(ws, "BudgetPerTicker")
    Dim lotSizeParamCol As Long: lotSizeParamCol = FindParamColumn(ws, "LotSize")

    Dim biasBpGlobal As Double: biasBpGlobal = GetParamDouble(ws, biasParamCol, 0#)
    Dim biasSlopeGlobal As Double: biasSlopeGlobal = GetParamDouble(ws, biasSlopeParamCol, 0.1)
    Dim gapSlopeGlobal As Double: gapSlopeGlobal = GetParamDouble(ws, gapSlopeParamCol, 0.2)
    Dim corrSlopeGlobal As Double: corrSlopeGlobal = GetParamDouble(ws, corrSlopeParamCol, 0.05)
    Dim gapBanPct As Double: gapBanPct = GetParamDouble(ws, gapBanParamCol, 3#)
    Dim tpParam As Double: tpParam = GetParamDouble(ws, tpParamCol, 0.15)
    Dim slParam As Double: slParam = GetParamDouble(ws, slParamCol, 0.1)
    Dim trailParam As Double: trailParam = GetParamDouble(ws, trailParamCol, 0.1)
    Dim budgetPerTicker As Double: budgetPerTicker = GetParamDouble(ws, budgetParamCol, 1000000#)
    Dim lotSize As Double: lotSize = GetParamDouble(ws, lotSizeParamCol, 100#)
    If lotSize <= 0# Then lotSize = 1#

    Dim weeklySellRule As String: weeklySellRule = GetStrategyRule("weekly_sell", "allow")
    Dim jcrossRequireDown As Boolean: jcrossRequireDown = (GetStrategyRule("jcross_sell_require_nky_down", "1") = "1")
    Dim jcrossMinGap As Double: jcrossMinGap = GetStrategyRuleDouble("jcross_sell_min_gap_bp", 20#)

    Dim r As Long
    For r = DASH2_DATA_START To lastRow
        Dim ticker As String: ticker = Trim$(CStr(ws.Cells(r, tickerCol).Value))
        If ticker = "" Then
            If CanWriteCell(ws, r, jthCol) Then ws.Cells(r, jthCol).Value = ""
            ClearRowDynamicCells ws, r, eBuyCol, eSellCol, sideCol, tpEffCol, slEffCol, trailEffCol, qtyCol, volTagCol
            GoTo ContinueLoop
        End If

        Dim baseJ As Double: baseJ = ToDouble(ws.Cells(r, jthBaseCol).Value, 0#)
        If baseJ = 0# And Trim$(CStr(ws.Cells(r, jthBaseCol).Value)) = "" Then
            If CanWriteCell(ws, r, jthCol) Then ws.Cells(r, jthCol).Value = ""
            ClearRowDynamicCells ws, r, eBuyCol, eSellCol, sideCol, tpEffCol, slEffCol, trailEffCol, qtyCol, volTagCol
            GoTo ContinueLoop
        End If

        Dim tickerVal As String
        If tickerCol > 0 Then tickerVal = Trim$(CStr(ws.Cells(r, tickerCol).Value))
        Dim sessionVal As String
        If sessionCol > 0 Then sessionVal = Trim$(CStr(ws.Cells(r, sessionCol).Value))
        Dim driverVal As String
        driverVal = DRIVER_NKY
        If trendDriverCol > 0 Then
            Dim driverRaw As String
            driverRaw = NormalizeDriverName(CStr(ws.Cells(r, trendDriverCol).Value))
            If Len(driverRaw) > 0 Then
                driverVal = driverRaw
            End If
        End If
        Dim gapVal As Double: gapVal = ToDouble(ws.Cells(r, gapCol).Value, 0#)
        Dim corrVal As Double: corrVal = ResolveDriverCorrelation(ws, r, driverVal, corrCol, corrTopixCol)
        Dim gapAbsPct As Double: gapAbsPct = Abs(gapVal) / 100#

        If gapBanPct > 0# And gapAbsPct > gapBanPct Then
            If CanWriteCell(ws, r, jthCol) Then ws.Cells(r, jthCol).Value = "BAN"
            ClearRowDynamicCells ws, r, eBuyCol, eSellCol, sideCol, tpEffCol, slEffCol, trailEffCol, qtyCol, volTagCol
            If volTagCol > 0 Then ws.Cells(r, volTagCol).Value = "BAN"
            GoTo ContinueLoop
        End If

        Dim biasSlopeVal As Double: biasSlopeVal = RowWithFallback(ws, r, biasSlopeRowCol, biasSlopeGlobal)
        Dim gapSlopeVal As Double: gapSlopeVal = RowWithFallback(ws, r, gapSlopeRowCol, gapSlopeGlobal)
        Dim corrSlopeVal As Double: corrSlopeVal = RowWithFallback(ws, r, corrSlopeRowCol, corrSlopeGlobal)

        Dim adjJth As Double
        adjJth = baseJ + biasSlopeVal * (biasBpGlobal / 100#) + gapSlopeVal * gapAbsPct + corrSlopeVal * corrVal * (biasBpGlobal / 100#)
        If CanWriteCell(ws, r, jthCol) Then ws.Cells(r, jthCol).Value = adjJth

        Dim vwapVal As Double: vwapVal = ToDouble(ws.Cells(r, vwapCol).Value, 0#)
        Dim prevVal As Double: prevVal = ToDouble(ws.Cells(r, prevCol).Value, 0#)
        Dim lastVal As Double: lastVal = 0#
        If lastCol > 0 Then
            lastVal = ToDouble(ws.Cells(r, lastCol).Value, 0#)
        Else
            lastVal = 0#
        End If
        Dim basePrice As Double: basePrice = DetermineBasePrice(vwapVal, prevVal)
        If basePrice <= 0# Then basePrice = lastVal

        If basePrice > 0# Then
            If CanWriteCell(ws, r, eBuyCol) Then ws.Cells(r, eBuyCol).Value = basePrice - 0.001 * Abs(adjJth) * basePrice
            If CanWriteCell(ws, r, eSellCol) Then ws.Cells(r, eSellCol).Value = basePrice + 0.001 * Abs(adjJth) * basePrice
        Else
            ClearRowDynamicCells ws, r, eBuyCol, eSellCol
        End If

        Dim jVal As Double: jVal = ToDouble(ws.Cells(r, jCol).Value, 0#)
        Dim entrySide As String
        If jVal < 0# Then
            entrySide = "BUY"
        ElseIf jVal > 0# Then
            entrySide = "SELL"
        Else
            entrySide = ""
        End If
        If sideCol > 0 Then
            If entrySide = "" Then
                If CanWriteCell(ws, r, sideCol) Then ws.Cells(r, sideCol).ClearContents
            Else
                If CanWriteCell(ws, r, sideCol) Then ws.Cells(r, sideCol).Value = entrySide
            End If
        End If

        Dim ratioVal As Double
        If adjJth <> 0# Then
            ratioVal = Abs(jVal) / Abs(adjJth)
        Else
            ratioVal = 0#
        End If
        Dim driverDayState As String: driverDayState = GetDriverTrendDay(driverVal)
        Dim driverWindowState As String: driverWindowState = GetDriverTrendWindow(driverVal)
        Dim allowedSideState As String: allowedSideState = GetDriverAllowedSide(driverVal)
        If driverDayCol > 0 Then ws.Cells(r, driverDayCol).Value = driverDayState
        If driverWindowCol2 > 0 Then ws.Cells(r, driverWindowCol2).Value = driverWindowState
        If driverAllowedCol > 0 Then ws.Cells(r, driverAllowedCol).Value = allowedSideState
        Dim trendWindowVal As String
        If trendWindowCol > 0 Then
            trendWindowVal = Trim$(CStr(ws.Cells(r, trendWindowCol).Value))
        End If
        If Len(trendWindowVal) = 0 Then
            trendWindowVal = driverWindowState
        End If
        Dim bbRisk As BbRiskLevel
        bbRisk = EvaluateBbRisk(tickerVal, sessionVal, trendWindowVal, ratioVal)
        Dim bbKey As String
        If Len(tickerVal) > 0 And Len(sessionVal) > 0 Then
            bbKey = GetBbBlockKey(tickerVal, sessionVal)
        End If
        Dim bbActive As Boolean
        If Len(bbKey) > 0 Then
            bbActive = IsBbBlockActive(bbKey)
        End If

        Dim blockReason As String: blockReason = ""
        Dim batchKindVal As String
        If batchKindCol > 0 Then batchKindVal = Trim$(CStr(ws.Cells(r, batchKindCol).Value))
        Dim signalModeVal As String
        If modeCol > 0 Then signalModeVal = Trim$(CStr(ws.Cells(r, modeCol).Value))
        Dim filterSide As String
        filterSide = allowedSideState
        If Len(filterSide) = 0 Then filterSide = "BOTH"
        If entrySide <> "" And filterSide <> "BOTH" Then
            If StrComp(entrySide, filterSide, vbTextCompare) <> 0 Then
                blockReason = "BLOCKED_DIR_" & driverVal & "_" & UCase$(filterSide)
            End If
        End If
        Dim isWeekend As Boolean
        isWeekend = (StrComp(batchKindVal, "weekend", vbTextCompare) = 0)
        If blockReason = "" And entrySide = "SELL" And weeklySellRule = "disable" And isWeekend Then
            blockReason = "BLOCKED_WEEKEND_SELL"
        End If
        If blockReason = "" And entrySide = "SELL" And StrComp(signalModeVal, "j-cross", vbTextCompare) = 0 Then
            If jcrossRequireDown Then
                If Not DriverHasDownTrend(driverVal) Then
                    blockReason = "BLOCKED_JCROSS_TREND_" & driverVal
                End If
            End If
            If blockReason = "" Then
                Dim gapGate As Double
                If gapCol > 0 Then gapGate = ToDouble(ws.Cells(r, gapCol).Value, 0#) Else gapGate = 0#
                If Abs(gapGate) < jcrossMinGap Then
                    blockReason = "BLOCKED_JCROSS_GAP"
                End If
            End If
        End If
        If blockReason = "" And ratioVal >= SPIKE_RATIO_THRESHOLD Then
            blockReason = "BLOCKED_SPIKE"
            AppendSpikeEvent tickerVal, sessionVal, ratioVal
        End If

        If bbRisk = bbRiskBlock And Len(bbKey) > 0 Then
            ActivateBbBlock bbKey
            bbActive = True
        End If

        If blockReason = "" Then
            If bbActive Then
                blockReason = "BLOCKED_BB"
            ElseIf bbRisk = bbRiskWarn Then
                If entryStatusCol > 0 Then ws.Cells(r, entryStatusCol).Value = "WARN_BB"
            End If
        End If

        ApplyDirectionHighlight ws, r, entrySide, allowedSideState
        If blockReason <> "" Then
            If entryStatusCol > 0 Then ws.Cells(r, entryStatusCol).Value = blockReason
            Dim shouldDeselect As Boolean: shouldDeselect = True
            If Left$(blockReason, 12) = "BLOCKED_DIR_" Or Left$(blockReason, 21) = "BLOCKED_JCROSS_TREND_" Or blockReason = "BLOCKED_BB" Then
                shouldDeselect = False
            End If
            If shouldDeselect And selCol > 0 Then ws.Cells(r, selCol).Value = 0
            ClearRowDynamicCells ws, r, eBuyCol, eSellCol, tpEffCol, slEffCol, trailEffCol, qtyCol
            GoTo ContinueLoop
        Else
            If entryStatusCol > 0 Then
                Dim curStatus As String: curStatus = CStr(ws.Cells(r, entryStatusCol).Value)
                If Left$(curStatus, 7) = "BLOCKED" Then
                    ws.Cells(r, entryStatusCol).ClearContents
                    If selCol > 0 And ws.Cells(r, selCol).Value = 0 Then
                        ws.Cells(r, selCol).Value = 1
                    End If
                ElseIf StrComp(curStatus, "WARN_BB", vbTextCompare) = 0 And bbRisk <> bbRiskWarn Then
                    ws.Cells(r, entryStatusCol).ClearContents
                End If
            End If
        End If

        Dim priceForQty As Double: priceForQty = lastVal
        If priceForQty <= 0# Then priceForQty = basePrice
        Dim qty As Double
        Dim effBudget As Double
        Dim budgetFactor As Double
        If budgetFactorCol > 0 Then
            budgetFactor = ToDouble(ws.Cells(r, budgetFactorCol).Value, 1#)
            If budgetFactor <= 0# Then budgetFactor = 1#
        Else
            budgetFactor = 1#
        End If
        effBudget = budgetPerTicker * budgetFactor
        If priceForQty > 0# Then
            qty = Int(effBudget / priceForQty / lotSize) * lotSize
            If qty < 0# Then qty = 0#
        Else
            qty = 0#
        End If
        If qtyCol > 0 Then
            If CanWriteCell(ws, r, qtyCol) Then
                If qty > 0# Then
                    ws.Cells(r, qtyCol).Value = qty
                Else
                    ws.Cells(r, qtyCol).ClearContents
                End If
            End If
        End If

        Dim atrVal As Double: atrVal = ToDouble(ws.Cells(r, atrCol).Value, 0#)
        Dim volFactor As Double: volFactor = ComputeVolFactor(gapAbsPct, Abs(jVal), atrVal)
        If volTagCol > 0 Then ws.Cells(r, volTagCol).Value = ClassifyVolFactor(volFactor)

        Dim tpBase As Double: tpBase = ResolveRowBase(ws, r, tpRowCol, tpParam)
        Dim slBase As Double: slBase = ResolveRowBase(ws, r, slRowCol, slParam)
        Dim trailBase As Double: trailBase = ResolveRowBase(ws, r, trailRowCol, trailParam)

        Dim tpEff As Double: tpEff = Round(tpBase * volFactor, 4)
        Dim slEff As Double: slEff = Round(slBase * volFactor, 4)
        Dim trailEff As Double: trailEff = Round(trailBase * volFactor, 4)

        If tpRowCol > 0 Then ws.Cells(r, tpRowCol).Value = tpBase
        If slRowCol > 0 Then ws.Cells(r, slRowCol).Value = slBase
        If trailRowCol > 0 Then ws.Cells(r, trailRowCol).Value = trailBase
        If tpEffCol > 0 And CanWriteCell(ws, r, tpEffCol) Then ws.Cells(r, tpEffCol).Value = tpEff
        If slEffCol > 0 And CanWriteCell(ws, r, slEffCol) Then ws.Cells(r, slEffCol).Value = slEff
        If trailEffCol > 0 And CanWriteCell(ws, r, trailEffCol) Then ws.Cells(r, trailEffCol).Value = trailEff

ContinueLoop:
    Next r

End Sub



Public Sub PreplaceOrdersV2()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim selCol As Long: selCol = FindColumn(ws, DASH2_HEADER_ROW, "Selected")
    Dim eBuyCol As Long: eBuyCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryBuyPx")
    Dim eSellCol As Long: eSellCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySellPx")
    Dim qtyCol As Long: qtyCol = FindColumn(ws, DASH2_HEADER_ROW, "OrderQtyPlan")
    Dim tpCol As Long: tpCol = FindColumn(ws, DASH2_HEADER_ROW, "TP_price")
    Dim slCol As Long: slCol = FindColumn(ws, DASH2_HEADER_ROW, "SL_price")
    Dim modeCol As Long: modeCol = FindColumn(ws, DASH2_HEADER_ROW, "SignalMode")
    Dim sessionCol As Long: sessionCol = FindColumn(ws, DASH2_HEADER_ROW, "session")
    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderTickerJP())

    If selCol = 0 Or tickerCol = 0 Then Exit Sub

    Dim sh As Worksheet
    Set sh = EnsureOrdersSheet(ws)

    Dim slipDict As Object
    Set slipDict = LoadSlippageOverrides()

    Dim lastOrder As Long: lastOrder = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim idx As Long
    For idx = lastOrder To 2 Step -1
        If LCase$(CStr(sh.Cells(idx, 6).Value)) = "preplace" Then
            sh.Rows(idx).Delete
        End If
    Next idx

    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = DASH2_DATA_START To lastR
        If ws.Cells(r, selCol).Value = 1 Then
            Dim hasBuy As Boolean: hasBuy = False
            Dim hasSell As Boolean: hasSell = False
            Dim tmpVal As Variant
            If eBuyCol > 0 Then
                hasBuy = ws.Cells(r, eBuyCol).HasFormula
                If Not hasBuy Then
                    tmpVal = ws.Cells(r, eBuyCol).Value
                    hasBuy = IsNumeric(tmpVal) And tmpVal > 0
                End If
            End If
            If eSellCol > 0 Then
                hasSell = ws.Cells(r, eSellCol).HasFormula
                If Not hasSell Then
                    tmpVal = ws.Cells(r, eSellCol).Value
                    hasSell = IsNumeric(tmpVal) And tmpVal > 0
                End If
            End If
            If hasBuy Or hasSell Then
                Dim modeVal As String: modeVal = ""
                Dim sessionVal As String: sessionVal = ""
                If modeCol > 0 Then modeVal = CStr(ws.Cells(r, modeCol).Value)
                If sessionCol > 0 Then sessionVal = CStr(ws.Cells(r, sessionCol).Value)
                Dim noteExtra As String: noteExtra = ""
                Dim bufferFrac As Double: bufferFrac = 0
                Dim key As String: key = GetSlippageKey(sessionVal, modeVal)
                If Len(key) > 0 Then
                    If Not slipDict Is Nothing Then
                        If slipDict.Exists(key) Then
                            bufferFrac = CDbl(slipDict(key)) / 10000#
                            If bufferFrac < 0 Then bufferFrac = 0
                            If bufferFrac > 0.003 Then bufferFrac = 0.003
                            noteExtra = "buffer_bp=" & Format$(slipDict(key), "0.0")
                        End If
                    End If
                End If
                LogPreOrder ws, sh, r, hasBuy, hasSell, eBuyCol, eSellCol, qtyCol, tpCol, slCol, modeCol, sessionCol, tickerCol, bufferFrac, noteExtra
            End If
        End If
    Next r

End Sub



Private Function EnsureOrdersSheet(ByVal host As Worksheet) As Worksheet

    Dim sh As Worksheet

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("Orders")
    On Error GoTo 0

    If sh Is Nothing Then
        Dim anchor As Worksheet
        If host Is Nothing Then
            Set anchor = ThisWorkbook.Worksheets(1)
        Else
            Set anchor = host
        End If
        Set sh = ThisWorkbook.Worksheets.Add(After:=anchor)
        sh.Name = "Orders"
    End If

    With sh.Range("A1:R1")
        .Value = Array("ts", "ticker", "side", "price", "qty", "mode", "status", "note", "tp", "sl", "trail", "fill_ts", "fill_price", "fill_qty", "close_ts", "close_price", "pnl_bp", "source")
    End With

    Set EnsureOrdersSheet = sh

End Function

Private Function FindOrderRow(ByVal sh As Worksheet, ByVal ticker As String, ByVal side As String, ByVal statusFilters As Variant) As Long
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    Dim targetTicker As String: targetTicker = UCase$(Trim$(ticker))
    Dim targetSide As String: targetSide = UCase$(Trim$(side))
    For r = lastRow To 2 Step -1
        If UCase$(Trim$(CStr(sh.Cells(r, 2).Value))) = targetTicker Then
            If targetSide = "" Or UCase$(Trim$(CStr(sh.Cells(r, 3).Value))) = targetSide Then
                If IsEmpty(statusFilters) Then
                    FindOrderRow = r
                    Exit Function
                Else
                    Dim statusVal As String
                    statusVal = UCase$(Trim$(CStr(sh.Cells(r, 7).Value)))
                    Dim idx As Long
                    For idx = LBound(statusFilters) To UBound(statusFilters)
                        If UCase$(CStr(statusFilters(idx))) = statusVal Then
                            FindOrderRow = r
                            Exit Function
                        End If
                    Next idx
                End If
            End If
        End If
    Next r
    FindOrderRow = 0
End Function

Private Function AppendOrderRow(ByVal ticker As String, ByVal side As String, ByVal price As Double, ByVal qty As Double, ByVal mode As String, ByVal statusVal As String, ByVal note As String) As Long
    Dim sh As Worksheet: Set sh = EnsureOrdersSheet(Nothing)
    Dim rowIdx As Long
    rowIdx = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1
    sh.Cells(rowIdx, 1).Value = Now
    sh.Cells(rowIdx, 2).Value = ticker
    sh.Cells(rowIdx, 3).Value = side
    sh.Cells(rowIdx, 4).Value = price
    sh.Cells(rowIdx, 5).Value = qty
    sh.Cells(rowIdx, 6).Value = mode
    sh.Cells(rowIdx, 7).Value = statusVal
    sh.Cells(rowIdx, 8).Value = note
    AppendOrderRow = rowIdx
End Function

Public Sub LogOrderFilled(ByVal ticker As String, ByVal side As String, ByVal fillPrice As Double, ByVal qty As Double, Optional ByVal mode As String = "", Optional ByVal note As String = "")
    Dim sh As Worksheet: Set sh = EnsureOrdersSheet(Nothing)
    Dim targets As Variant: targets = Array("PENDING", "PREPLACE", "PREPLACE")
    Dim rowIdx As Long
    rowIdx = FindOrderRow(sh, ticker, side, targets)
    If rowIdx = 0 Then
        rowIdx = AppendOrderRow(ticker, side, fillPrice, qty, mode, "NEW", note)
    End If
    sh.Cells(rowIdx, 7).Value = "FILLED"
    sh.Cells(rowIdx, 12).Value = Now
    sh.Cells(rowIdx, 13).Value = fillPrice
    If qty > 0# Then sh.Cells(rowIdx, 14).Value = qty
    If Len(mode) > 0 Then sh.Cells(rowIdx, 6).Value = mode
    If Len(note) > 0 Then sh.Cells(rowIdx, 8).Value = note
End Sub

Public Sub LogOrderSettled(ByVal ticker As String, ByVal side As String, ByVal closePrice As Double, Optional ByVal qty As Double = 0#, Optional ByVal note As String = "")
    Dim sh As Worksheet: Set sh = EnsureOrdersSheet(Nothing)
    Dim rowIdx As Long
    rowIdx = FindOrderRow(sh, ticker, side, Array("FILLED", "RUNNING"))
    If rowIdx = 0 Then
        rowIdx = AppendOrderRow(ticker, side, closePrice, qty, "", "ADHOC_CLOSE", note)
    End If
    sh.Cells(rowIdx, 7).Value = "CLOSED"
    sh.Cells(rowIdx, 15).Value = Now
    sh.Cells(rowIdx, 16).Value = closePrice
    If qty > 0# Then sh.Cells(rowIdx, 14).Value = qty
    Dim fillPrice As Double: fillPrice = ToDouble(sh.Cells(rowIdx, 13).Value, 0#)
    Dim fillQty As Double: fillQty = ToDouble(sh.Cells(rowIdx, 14).Value, 0#)
    If fillQty = 0# And qty > 0# Then fillQty = qty
    If Len(note) > 0 Then sh.Cells(rowIdx, 8).Value = note
    If fillPrice > 0# And fillQty > 0# Then
        Dim direction As Double
        direction = IIf(UCase$(Trim$(side)) = "BUY", 1#, -1#)
        Dim pnlBp As Double
        pnlBp = direction * (closePrice - fillPrice) / fillPrice * 10000#
        sh.Cells(rowIdx, 17).Value = pnlBp
    End If
End Sub


Private Sub LogPreOrder(ByVal ws As Worksheet, ByVal orderSheet As Worksheet, ByVal rowIndex As Long, _
    ByVal includeBuy As Boolean, ByVal includeSell As Boolean, _
    ByVal eBuyCol As Long, ByVal eSellCol As Long, ByVal qtyCol As Long, _
    ByVal tpCol As Long, ByVal slCol As Long, ByVal modeCol As Long, _
    ByVal sessionCol As Long, ByVal tickerCol As Long, ByVal bufferFrac As Double, ByVal noteExtra As String)

    Dim sh As Worksheet
    Set sh = orderSheet
    If sh Is Nothing Then
        Set sh = EnsureOrdersSheet(ws)
    End If

    Dim sheetToken As String
    sheetToken = "'" & Replace(ws.Name, "'", "''") & "'!"

    Dim tickerRef As String
    If tickerCol > 0 Then
        tickerRef = sheetToken & ws.Cells(rowIndex, tickerCol).Address(True, True, xlA1)
    Else
        tickerRef = ""
    End If

    Dim qtyRef As String
    If qtyCol > 0 Then qtyRef = sheetToken & ws.Cells(rowIndex, qtyCol).Address(True, True, xlA1) Else qtyRef = ""

    Dim tpRef As String
    If tpCol > 0 Then tpRef = sheetToken & ws.Cells(rowIndex, tpCol).Address(True, True, xlA1) Else tpRef = ""

    Dim slRef As String
    If slCol > 0 Then slRef = sheetToken & ws.Cells(rowIndex, slCol).Address(True, True, xlA1) Else slRef = ""

    Dim sessionRef As String
    If sessionCol > 0 Then sessionRef = sheetToken & ws.Cells(rowIndex, sessionCol).Address(True, True, xlA1) Else sessionRef = ""

    Dim modeRef As String
    If modeCol > 0 Then modeRef = sheetToken & ws.Cells(rowIndex, modeCol).Address(True, True, xlA1) Else modeRef = ""

    Dim noteFormula As String
    noteFormula = BuildNoteFormula(sessionRef, modeRef, noteExtra)

    Dim nowStr As String: nowStr = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim nextRow As Long: nextRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1

    If includeBuy Then
        sh.Cells(nextRow, 1).Value = nowStr
        If tickerRef <> "" Then
            sh.Cells(nextRow, 2).Formula = "=" & tickerRef
        Else
            sh.Cells(nextRow, 2).Value = ""
        End If
        sh.Cells(nextRow, 3).Value = "BUY"
        If eBuyCol > 0 Then
            Dim buyRef As String
            buyRef = sheetToken & ws.Cells(rowIndex, eBuyCol).Address(True, True, xlA1)
            sh.Cells(nextRow, 4).Formula = BuildAdjustedPriceFormula(buyRef, bufferFrac, True)
        Else
            sh.Cells(nextRow, 4).Value = ""
        End If
        If qtyRef <> "" Then
            sh.Cells(nextRow, 5).Formula = "=" & qtyRef
        Else
            sh.Cells(nextRow, 5).Value = 0
        End If
        sh.Cells(nextRow, 6).Value = "preplace"
        sh.Cells(nextRow, 7).Value = "PENDING"
        If noteFormula <> "" Then
            sh.Cells(nextRow, 8).Formula = noteFormula
        Else
            sh.Cells(nextRow, 8).Value = ""
        End If
        If tpRef <> "" Then
            sh.Cells(nextRow, 9).Formula = "=" & tpRef
        Else
            sh.Cells(nextRow, 9).Value = ""
        End If
        If slRef <> "" Then
            sh.Cells(nextRow, 10).Formula = "=" & slRef
        Else
            sh.Cells(nextRow, 10).Value = ""
        End If
        sh.Cells(nextRow, 11).Value = ""
        nextRow = nextRow + 1
    End If

    If includeSell Then
        sh.Cells(nextRow, 1).Value = nowStr
        If tickerRef <> "" Then
            sh.Cells(nextRow, 2).Formula = "=" & tickerRef
        Else
            sh.Cells(nextRow, 2).Value = ""
        End If
        sh.Cells(nextRow, 3).Value = "SELL"
        If eSellCol > 0 Then
            Dim sellRef As String
            sellRef = sheetToken & ws.Cells(rowIndex, eSellCol).Address(True, True, xlA1)
            sh.Cells(nextRow, 4).Formula = BuildAdjustedPriceFormula(sellRef, bufferFrac, False)
        Else
            sh.Cells(nextRow, 4).Value = ""
        End If
        If qtyRef <> "" Then
            sh.Cells(nextRow, 5).Formula = "=" & qtyRef
        Else
            sh.Cells(nextRow, 5).Value = 0
        End If
        sh.Cells(nextRow, 6).Value = "preplace"
        sh.Cells(nextRow, 7).Value = "PENDING"
        If noteFormula <> "" Then
            sh.Cells(nextRow, 8).Formula = noteFormula
        Else
            sh.Cells(nextRow, 8).Value = ""
        End If
        If tpRef <> "" Then
            sh.Cells(nextRow, 9).Formula = "=" & tpRef
        Else
            sh.Cells(nextRow, 9).Value = ""
        End If
        If slRef <> "" Then
            sh.Cells(nextRow, 10).Formula = "=" & slRef
        Else
            sh.Cells(nextRow, 10).Value = ""
        End If
        sh.Cells(nextRow, 11).Value = ""
    End If

End Sub



Public Function PlaceOrderRSS(ByVal ticker As String, ByVal side As String, ByVal price As Double, ByVal qty As Long) As Boolean

    On Error GoTo Fail

    Dim evalStr As String

    evalStr = "RssOrder(\"" & ticker & " \ ",\"" & side & " \ "," & CStr(price) & "," & CStr(qty) & ")"

    Dim result As Variant

    result = Application.Evaluate("=" & evalStr)

    If IsError(result) Then GoTo Fail

    PlaceOrderRSS = True

    Exit Function

Fail:

    PlaceOrderRSS = False

End Function



' ----------------------------------------------------------------------------

' UI: Buttons/Status/Labels/Order

' ----------------------------------------------------------------------------

Public Sub SetupDashboardUIV2()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Dim wasProtected As Boolean
    On Error Resume Next
    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    If wasProtected Then ws.Unprotect
    On Error GoTo 0

    On Error GoTo SetupFail


    ' Status + direction indicators near the left
    With ws.Range("A3")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(230, 230, 230)
        .Value = "IDLE"
        .Name = "RunStatusV2"
    End With

    SetupTrendIndicatorCells ws
    UpdateTrendIndicators ws



    Dim shp As Shape

    For Each shp In ws.Shapes

        If shp.name Like "btn_*" Then shp.Delete

    Next shp



    ' Buttons (swap: Live on left, Demo on right)

    CreateButton ws, "btn_live_start", "Live Start", 3, 6, "AutoTraderAdvanced.StartLiveV2"    ' F3

    CreateButton ws, "btn_live_stop", "Live Stop", 3, 8, "AutoTraderAdvanced.StopLiveV2"      ' H3

    CreateButton ws, "btn_demo_start", "Demo Start", 3, 10, "AutoTraderAdvanced.StartDemoV2"    ' J3

    CreateButton ws, "btn_demo_stop", "Demo Stop", 3, 12, "AutoTraderAdvanced.StopDemoV2"      ' L3

    CreateButton ws, "btn_import", "Import Candidates", 3, 14, "AutoTraderAdvanced.ImportCandidatesV2"      ' N3

    ' ASCII-only captions to avoid encoding issues
    CreateButton ws, "btn_refresh_trend", "Recalc Trend", 3, 16, "AutoTraderAdvanced.RefreshTrendsV2"        ' P3

    CreateButton ws, "btn_clear_bb", "Clear BB Blocks", 3, 18, "AutoTraderAdvanced.ClearBBBlocks"          ' R3

    ApplyJapaneseLabelsV2 ws

    ReorderHeadersV2 ws

    UpdateTrendIndicators ws

    EnsureDashboardProtectionV2 ws
    Exit Sub

SetupFail:
    LogVbaEvent "SetupDashboardUIV2", "Err " & Err.Number & ": " & Err.Description
    On Error Resume Next
    EnsureDashboardProtectionV2 ws
    On Error GoTo 0
End Sub

Private Sub EnsureDashboardProtectionV2(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
               AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowFiltering:=True, AllowSorting:=True
    On Error GoTo 0
End Sub



Private Sub CreateButton(ByVal ws As Worksheet, ByVal name As String, ByVal caption As String, ByVal rowTop As Long, ByVal colLeft As Long, ByVal onAction As String)

    Dim leftPos As Double: leftPos = ws.Cells(rowTop, colLeft).Left

    Dim topPos As Double: topPos = ws.Cells(rowTop, colLeft).Top

    Dim width As Double: width = ws.Range(ws.Cells(rowTop, colLeft), ws.Cells(rowTop, colLeft + 3)).width - 6

    Dim height As Double: height = ws.Rows(rowTop).height - 2

    Dim btn As Shape

    Set btn = ws.Shapes.AddShape(MsoShapeRectangle, leftPos, topPos, width, height)

    On Error Resume Next

    btn.name = name

    On Error GoTo 0

    btn.TextFrame.Characters.Text = caption

    btn.TextFrame.HorizontalAlignment = MsoAlignCenter

    btn.TextFrame.VerticalAlignment = MsoAlignCenter

    btn.onAction = onAction

End Sub



Private Sub ApplyJapaneseLabelsV2(ByVal ws As Worksheet)

    Dim labels As Variant

    labels = Array("Ticker", "Name", "J_th_base", "J_th", "J", "PrevClose", "VWAP", "OrderQtyPlan", "Selected", _
                   "EntryBuyPx", "EntrySellPx", "EntrySide", "EntryStatus", "TP_price", "SL_price", "StopTrail", "SettleStatus", "BestBid", "BestAsk", "Gap_bp", "CorrNKY", "CorrTOPIX", _
                   "BiasSlope_row", "GapSlope_row", "CorrSlope_row", "TP_per_J_row", "SL_per_J_row", "Trail_per_J_row", "TP_per_J_eff", "SL_per_J_eff", "Trail_per_J_eff", "BudgetFactor_row", "VolatilityTag")

    Dim i As Long

    For i = 1 To UBound(labels) + 1

        ws.Cells(4, i).Value = labels(i - 1)

        On Error Resume Next

        ws.Cells(4, i).AddComment "Header description"

        On Error GoTo 0

    Next i

End Sub



Private Sub ReorderHeadersV2(ByVal ws As Worksheet)

    Dim order As Variant

    order = Array("Ticker", "Name", "J_th_base", "J_th", "J", "PrevClose", "VWAP", "OrderQtyPlan", "Selected", "EntryBuyPx", "EntrySellPx", "EntrySide", "EntryStatus", "TP_price", "SL_price", "StopTrail", "SettleStatus", "BestBid", "BestAsk", "Gap_bp", "CorrNKY", "CorrTOPIX", _
                  "BiasSlope_row", "GapSlope_row", "CorrSlope_row", "TP_per_J_row", "SL_per_J_row", "Trail_per_J_row", "TP_per_J_eff", "SL_per_J_eff", "Trail_per_J_eff", "BudgetFactor_row", "VolatilityTag")

    Dim c As Long

    For c = 1 To UBound(order) + 1

        ws.Cells(DASH2_HEADER_ROW, c).Value = CStr(order(c - 1))

    Next c

End Sub



Public Sub StartDemoV2(): UpdateStatusV2 "DEMO_RUNNING": End Sub

Public Sub StopDemoV2():  UpdateStatusV2 "IDLE": End Sub

Public Sub StartLiveV2(): UpdateStatusV2 "LIVE_RUNNING": End Sub

Public Sub StopLiveV2():  UpdateStatusV2 "IDLE": End Sub

Private Sub UpdateStatusV2(ByVal mode As String)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    Dim statusCell As Range
    Set statusCell = ws.Range("A3")

    On Error Resume Next
    ws.Parent.Names.Add Name:="RunStatusV2", RefersTo:=statusCell
    On Error GoTo 0

    statusCell.ClearContents

    With statusCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14

        Select Case mode
            Case "DEMO_RUNNING": .Interior.Color = RGB(220, 240, 255)
            Case "LIVE_RUNNING": .Interior.Color = RGB(255, 230, 230)
            Case Else: .Interior.Color = RGB(230, 230, 230)
        End Select
        .Value = mode
    End With

End Sub



Public Sub ImportCandidatesV2()

    Dim ws As Worksheet
    Dim path As String
    Dim f As Integer
    Dim raw As String
    Dim lines As Variant
    Dim line As Variant
    Dim importedCount As Long: importedCount = 0
    Dim r As Long: r = DASH2_DATA_START

    On Error GoTo ImportErr

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    Dim wasProtected As Boolean
    On Error Resume Next
    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    If wasProtected Then ws.Unprotect
    On Error GoTo 0

    LogVbaEvent "ImportCandidatesV2", "start workbook_path=" & ThisWorkbook.path & " dash=" & DASH2_SHEET & " protected=" & CStr(wasProtected)

    path = ThisWorkbook.path & "\output\excel\candidates_nextday.csv"

    On Error Resume Next

    If Len(Dir$(path)) = 0 Then

        Dim dt As String: dt = Format$(Date, "yyyymmdd")

        path = ThisWorkbook.path & "\output\excel\weekly_candidates_" & dt & ".csv"

    End If

    On Error GoTo 0

    ' Fallback to fixed workspace path (helps when workbook is opened from a different folder)
    If Len(Dir$(path)) = 0 Then
        path = "C:\AI\asagake\output\excel\candidates_nextday.csv"
    End If

    If Len(Dir$(path)) = 0 Then
        LogVbaEvent "ImportCandidatesV2", "candidate_csv_not_found path_try=" & path & " workbook_path=" & ThisWorkbook.path
        If wasProtected Then ProtectDashboardV2 ws
        Exit Sub
    End If

    ' Read full file and split by LF/CRLF. pandas emits LF-only by default,
    ' which can make Line Input treat the file as a single "line".
    f = FreeFile
    Open path For Binary As #f
    raw = String$(LOF(f), vbNullChar)
    If LOF(f) > 0 Then Get #f, , raw
    Close #f
    f = 0

    raw = Replace$(raw, vbCrLf, vbLf)
    raw = Replace$(raw, vbCr, vbLf)
    lines = Split(raw, vbLf)

    Dim selCol As Long: selCol = FindColumn(ws, DASH2_HEADER_ROW, "Selected")

    Dim jtbCol As Long: jtbCol = FindColumn(ws, DASH2_HEADER_ROW, "J_th_base")

    Dim colPf As Long: colPf = FindColumn(ws, DASH2_HEADER_ROW, "ForwardPfEff")

    Dim colCi As Long: colCi = FindColumn(ws, DASH2_HEADER_ROW, "WinCiLow")

    Dim colTr As Long: colTr = FindColumn(ws, DASH2_HEADER_ROW, "ForwardTrades")

    Dim colExp As Long: colExp = FindColumn(ws, DASH2_HEADER_ROW, "ExpBp")

    Dim colAtr As Long: colAtr = FindColumn(ws, DASH2_HEADER_ROW, "ATR_n")
    Dim colBiasSlope As Long: colBiasSlope = FindColumn(ws, DASH2_HEADER_ROW, "BiasSlope_row")
    Dim colGapSlope As Long: colGapSlope = FindColumn(ws, DASH2_HEADER_ROW, "GapSlope_row")
    Dim colCorrSlope As Long: colCorrSlope = FindColumn(ws, DASH2_HEADER_ROW, "CorrSlope_row")
    Dim colTpRow As Long: colTpRow = FindColumn(ws, DASH2_HEADER_ROW, "TP_per_J_row")
    Dim colSlRow As Long: colSlRow = FindColumn(ws, DASH2_HEADER_ROW, "SL_per_J_row")
    Dim colTrailRow As Long: colTrailRow = FindColumn(ws, DASH2_HEADER_ROW, "Trail_per_J_row")
    Dim colBudgetFactor As Long: colBudgetFactor = FindColumn(ws, DASH2_HEADER_ROW, "BudgetFactor_row")
    Dim colTrendDriver As Long: colTrendDriver = FindColumn(ws, DASH2_HEADER_ROW, "trend_driver")
    Dim colTrendWindow As Long: colTrendWindow = FindColumn(ws, DASH2_HEADER_ROW, "trend_window")
    Dim colTrendBp As Long: colTrendBp = FindColumn(ws, DASH2_HEADER_ROW, "trend_bp_th")
    Dim colTrendPolicy As Long: colTrendPolicy = FindColumn(ws, DASH2_HEADER_ROW, "trend_allowed_policy")
    Dim corrCol As Long: corrCol = FindColumn(ws, DASH2_HEADER_ROW, "CorrNKY")
    Dim corrTopixCol As Long: corrTopixCol = FindColumn(ws, DASH2_HEADER_ROW, "CorrTOPIX")

    Dim colTpk As Long: colTpk = FindColumn(ws, DASH2_HEADER_ROW, "TPk")

    Dim colSlk As Long: colSlk = FindColumn(ws, DASH2_HEADER_ROW, "SLk")

    Dim colMode As Long: colMode = FindColumn(ws, DASH2_HEADER_ROW, "SignalMode")

    Dim colSession As Long: colSession = FindColumn(ws, DASH2_HEADER_ROW, "session")

    Dim colPlan As Long: colPlan = FindColumn(ws, DASH2_HEADER_ROW, "plan_tag")
    Dim colBatchKind As Long: colBatchKind = FindColumn(ws, DASH2_HEADER_ROW, "BatchKind")

    Dim maxExisting As Long: maxExisting = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim first As Boolean: first = True

    Dim idxTicker As Long, idxJtb As Long, idxPf As Long, idxCi As Long, idxTrades As Long, idxExpBp As Long

    Dim idxAtr As Long, idxTpk As Long, idxSlk As Long, idxMode As Long, idxSession As Long, idxPlan As Long, idxBatchKind As Long
    Dim idxBiasSlope As Long, idxGapSlope As Long, idxCorrSlope As Long
    Dim idxTpRow As Long, idxSlRow As Long, idxTrailRow As Long, idxBudgetFactor As Long
    Dim idxCorrNky As Long, idxCorrTopix As Long
    Dim idxTrendDriver As Long, idxTrendWindow As Long, idxTrendBp As Long, idxTrendPolicy As Long

    idxTicker = -1: idxJtb = -1: idxPf = -1: idxCi = -1: idxTrades = -1: idxExpBp = -1

    idxAtr = -1: idxTpk = -1: idxSlk = -1: idxMode = -1: idxSession = -1: idxPlan = -1: idxBatchKind = -1
    idxBiasSlope = -1: idxGapSlope = -1: idxCorrSlope = -1
    idxTpRow = -1: idxSlRow = -1: idxTrailRow = -1: idxBudgetFactor = -1
    idxCorrNky = -1: idxCorrTopix = -1
    idxTrendDriver = -1: idxTrendWindow = -1: idxTrendBp = -1: idxTrendPolicy = -1

    Dim hdr As Variant
    For Each line In lines
        Dim lineText As String
        lineText = Trim$(CStr(line))
        If Len(lineText) = 0 Then GoTo NextLine

        If first Then

            hdr = ParseCsvLine(lineText)
            Dim i As Long

            For i = LBound(hdr) To UBound(hdr)

                Dim h As String: h = LCase$(Trim$(hdr(i)))
                h = Application.WorksheetFunction.Clean(h)
                Do While Len(h) > 0
                    Dim chCode As Long: chCode = AscW(Left$(h, 1))
                    If chCode >= 48 And chCode <= 122 Then Exit Do
                    h = Mid$(h, 2)
                Loop
                If Len(h) > 0 Then
                    If AscW(Left$(h, 1)) = &HFEFF Then
                        h = Mid$(h, 2)
                    End If
                End If
                If h = "ticker" Then idxTicker = i
                If h = "code" And idxTicker = -1 Then idxTicker = i

                If h = "j_th_base" Or h = "j_th" Then idxJtb = i

                If h = "forward_pf_eff" Or h = "pf" Then idxPf = i

                If h = "forward_win_ci_low" Or h = "win_ci_low" Then idxCi = i

                If h = "forward_trades" Or h = "trades" Then idxTrades = i

                If h = "exp_bp" Or h = "expected_bp" Or h = "forward_exp_bp" Or h = "forward_exp_boot_mean" Then idxExpBp = i
                If h = "atr_n" Then idxAtr = i

                If h = "tpk" Then idxTpk = i

                If h = "slk" Then idxSlk = i

                If h = "signalmode" Or h = "signal_mode" Then idxMode = i

                If h = "session" Then idxSession = i

                If h = "plan_tag" Then idxPlan = i
                If h = "batchkind" Or h = "batch_kind" Then idxBatchKind = i
                Dim hCore As String: hCore = h
                If Right$(hCore, 4) = "_row" Then hCore = Left$(hCore, Len(hCore) - 4)

                If h = "bias_slope" Or h = "biasslope" Or hCore = "biasslope" Then idxBiasSlope = i
                If h = "gap_slope" Or h = "gapslope" Or hCore = "gapslope" Then idxGapSlope = i
                If h = "corr_slope" Or h = "corrslope" Or hCore = "corrslope" Then idxCorrSlope = i
                If h = "tp_per_j" Or h = "tp_per_j_row" Or hCore = "tp_per_j" Then idxTpRow = i
                If h = "sl_per_j" Or h = "sl_per_j_row" Or hCore = "sl_per_j" Then idxSlRow = i
                If h = "trail_per_j" Or h = "trail_per_j_row" Or hCore = "trail_per_j" Then idxTrailRow = i
                If h = "budgetfactor_row" Or h = "budget_factor_row" Or hCore = "budgetfactor" Then idxBudgetFactor = i
                If h = "trend_driver" Then idxTrendDriver = i
                If h = "trend_window" Then idxTrendWindow = i
                If h = "trend_bp_th" Then idxTrendBp = i
                If h = "trend_allowed_policy" Then idxTrendPolicy = i
                If h = "corrnky" Or h = "corr_nky" Or hCore = "corrnky" Then idxCorrNky = i
                If h = "corrtopix" Or h = "corr_topix" Or hCore = "corrtopix" Then idxCorrTopix = i

            Next i

            first = False

        Else

            Dim parts As Variant: parts = ParseCsvLine(lineText)
            If idxTicker >= 0 And idxTicker <= UBound(parts) Then

                Dim tkr As String: tkr = Trim$(parts(idxTicker))
                If Len(tkr) > 1 And Left$(tkr, 1) = """" And Right$(tkr, 1) = """" Then
                    tkr = Mid$(tkr, 2, Len(tkr) - 2)
                End If
                tkr = Replace$(tkr, """", "")
                If Len(tkr) > 0 Then

                    ' Be tolerant of per-row write errors so that
                    ' one bad value does not abort the whole import.
                    On Error Resume Next

                    ws.Cells(r, 1).Value = tkr
                    If Err.Number <> 0 Then
                        LogVbaEvent "ImportCandidatesV2", "write_failed row=" & CStr(r) & " col=1 err=" & CStr(Err.Number) & " " & Err.Description & " ticker=" & tkr
                        Err.Clear
                        On Error GoTo ImportErr
                        GoTo ImportFinalize
                    End If

                    If selCol > 0 Then ws.Cells(r, selCol).Value = 1

                    If idxJtb >= 0 And idxJtb <= UBound(parts) And jtbCol > 0 Then

                        If Len(Trim$(parts(idxJtb))) > 0 Then ws.Cells(r, jtbCol).Value = Trim$(parts(idxJtb))

                    End If

                    If idxPf >= 0 And idxPf <= UBound(parts) And colPf > 0 Then ws.Cells(r, colPf).Value = Trim$(parts(idxPf))

                    If idxCi >= 0 And idxCi <= UBound(parts) And colCi > 0 Then ws.Cells(r, colCi).Value = Trim$(parts(idxCi))

                    If idxTrades >= 0 And idxTrades <= UBound(parts) And colTr > 0 Then ws.Cells(r, colTr).Value = Trim$(parts(idxTrades))

                    If idxExpBp >= 0 And idxExpBp <= UBound(parts) And colExp > 0 Then ws.Cells(r, colExp).Value = Trim$(parts(idxExpBp))

                    If idxAtr >= 0 And idxAtr <= UBound(parts) And colAtr > 0 Then

                        If Len(Trim$(parts(idxAtr))) > 0 Then

                            ws.Cells(r, colAtr).Value = Trim$(parts(idxAtr))

                        ElseIf ws.Cells(r, colAtr).Value = "" Then

                            ws.Cells(r, colAtr).Value = 2

                        End If

                    ElseIf colAtr > 0 And ws.Cells(r, colAtr).Value = "" Then

                        ws.Cells(r, colAtr).Value = 2

                    End If

                    If idxTpk >= 0 And idxTpk <= UBound(parts) And colTpk > 0 Then ws.Cells(r, colTpk).Value = Trim$(parts(idxTpk))

                    If idxSlk >= 0 And idxSlk <= UBound(parts) And colSlk > 0 Then ws.Cells(r, colSlk).Value = Trim$(parts(idxSlk))

                    If idxMode >= 0 And idxMode <= UBound(parts) And colMode > 0 Then ws.Cells(r, colMode).Value = Trim$(parts(idxMode))

                    If idxSession >= 0 And idxSession <= UBound(parts) And colSession > 0 Then ws.Cells(r, colSession).Value = Trim$(parts(idxSession))

                    If idxPlan >= 0 And idxPlan <= UBound(parts) And colPlan > 0 Then ws.Cells(r, colPlan).Value = Trim$(parts(idxPlan))
                    If idxBatchKind >= 0 And idxBatchKind <= UBound(parts) And colBatchKind > 0 Then ws.Cells(r, colBatchKind).Value = Trim$(parts(idxBatchKind))

                    If idxBiasSlope >= 0 And idxBiasSlope <= UBound(parts) And colBiasSlope > 0 Then
                        ws.Cells(r, colBiasSlope).Value = ToDouble(Trim$(parts(idxBiasSlope)), 0#)
                    End If
                    If idxGapSlope >= 0 And idxGapSlope <= UBound(parts) And colGapSlope > 0 Then
                        ws.Cells(r, colGapSlope).Value = ToDouble(Trim$(parts(idxGapSlope)), 0#)
                    End If
                    If idxCorrSlope >= 0 And idxCorrSlope <= UBound(parts) And colCorrSlope > 0 Then
                        ws.Cells(r, colCorrSlope).Value = ToDouble(Trim$(parts(idxCorrSlope)), 0#)
                    End If
                    If idxCorrNky >= 0 And idxCorrNky <= UBound(parts) And corrCol > 0 Then
                        ws.Cells(r, corrCol).Value = ToDouble(Trim$(parts(idxCorrNky)), 0#)
                    End If
                    If idxCorrTopix >= 0 And idxCorrTopix <= UBound(parts) And corrTopixCol > 0 Then
                        ws.Cells(r, corrTopixCol).Value = ToDouble(Trim$(parts(idxCorrTopix)), 0#)
                    End If
                    If idxTpRow >= 0 And idxTpRow <= UBound(parts) And colTpRow > 0 Then
                        ws.Cells(r, colTpRow).Value = ToDouble(Trim$(parts(idxTpRow)), 0#)
                    End If
                    If idxSlRow >= 0 And idxSlRow <= UBound(parts) And colSlRow > 0 Then
                        ws.Cells(r, colSlRow).Value = ToDouble(Trim$(parts(idxSlRow)), 0#)
                    End If
                    If idxTrailRow >= 0 And idxTrailRow <= UBound(parts) And colTrailRow > 0 Then
                        ws.Cells(r, colTrailRow).Value = ToDouble(Trim$(parts(idxTrailRow)), 0#)
                    End If
                    If idxBudgetFactor >= 0 And idxBudgetFactor <= UBound(parts) And colBudgetFactor > 0 Then
                        ws.Cells(r, colBudgetFactor).Value = ToDouble(Trim$(parts(idxBudgetFactor)), 1#)
                    End If
                    If idxTrendDriver >= 0 And idxTrendDriver <= UBound(parts) And colTrendDriver > 0 Then
                        ws.Cells(r, colTrendDriver).Value = Trim$(parts(idxTrendDriver))
                    End If
                    If idxTrendWindow >= 0 And idxTrendWindow <= UBound(parts) And colTrendWindow > 0 Then
                        ws.Cells(r, colTrendWindow).Value = Trim$(parts(idxTrendWindow))
                    End If
                    If idxTrendBp >= 0 And idxTrendBp <= UBound(parts) And colTrendBp > 0 Then
                        ws.Cells(r, colTrendBp).Value = ToDouble(Trim$(parts(idxTrendBp)), 0#)
                    End If
                    If idxTrendPolicy >= 0 And idxTrendPolicy <= UBound(parts) And colTrendPolicy > 0 Then
                        ws.Cells(r, colTrendPolicy).Value = Trim$(parts(idxTrendPolicy))
                    End If

                    On Error GoTo ImportErr

                    r = r + 1
                    importedCount = importedCount + 1

                End If

            End If

        End If

NextLine:
    Next line

    GoTo ImportFinalize

ImportFinalize:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0

    LogVbaEvent "ImportCandidatesV2", "done imported=" & CStr(importedCount) & " path=" & path

    Dim clearRow As Long

    On Error Resume Next
    ApplyDynamicSignalsV2
    On Error GoTo 0

    ' If we imported nothing, do not clear the sheet (keeps last known candidates visible).
    If importedCount > 0 And maxExisting >= DASH2_DATA_START And r <= maxExisting Then

        For clearRow = r To maxExisting

            ws.Cells(clearRow, 1).ClearContents

            If selCol > 0 Then ws.Cells(clearRow, selCol).ClearContents

            If jtbCol > 0 Then ws.Cells(clearRow, jtbCol).ClearContents

            If colPf > 0 Then ws.Cells(clearRow, colPf).ClearContents

            If colCi > 0 Then ws.Cells(clearRow, colCi).ClearContents

            If colTr > 0 Then ws.Cells(clearRow, colTr).ClearContents

            If colExp > 0 Then ws.Cells(clearRow, colExp).ClearContents

            If colAtr > 0 Then ws.Cells(clearRow, colAtr).ClearContents

            If colTpk > 0 Then ws.Cells(clearRow, colTpk).ClearContents

            If colSlk > 0 Then ws.Cells(clearRow, colSlk).ClearContents

            If colMode > 0 Then ws.Cells(clearRow, colMode).ClearContents

            If colSession > 0 Then ws.Cells(clearRow, colSession).ClearContents

            If colPlan > 0 Then ws.Cells(clearRow, colPlan).ClearContents

            If corrCol > 0 Then ws.Cells(clearRow, corrCol).ClearContents
            If corrTopixCol > 0 Then ws.Cells(clearRow, corrTopixCol).ClearContents

            If colBiasSlope > 0 Then ws.Cells(clearRow, colBiasSlope).ClearContents

            If colGapSlope > 0 Then ws.Cells(clearRow, colGapSlope).ClearContents

            If colCorrSlope > 0 Then ws.Cells(clearRow, colCorrSlope).ClearContents

            If colTpRow > 0 Then ws.Cells(clearRow, colTpRow).ClearContents

            If colSlRow > 0 Then ws.Cells(clearRow, colSlRow).ClearContents

            If colTrailRow > 0 Then ws.Cells(clearRow, colTrailRow).ClearContents
            If colBudgetFactor > 0 Then ws.Cells(clearRow, colBudgetFactor).ClearContents

            If colTrendDriver > 0 Then ws.Cells(clearRow, colTrendDriver).ClearContents

            If colTrendWindow > 0 Then ws.Cells(clearRow, colTrendWindow).ClearContents

            If colTrendBp > 0 Then ws.Cells(clearRow, colTrendBp).ClearContents

            If colTrendPolicy > 0 Then ws.Cells(clearRow, colTrendPolicy).ClearContents

        Next clearRow

    End If

    EnsureParamFormulas ws
    InstallRealtimeFormulasV2
    RefreshTrendsV2

    If wasProtected Then
        ProtectDashboardV2 ws
    End If
    Exit Sub

ImportErr:
    LogVbaEvent "ImportCandidatesV2", "Err " & Err.Number & ": " & Err.Description & " path=" & path
    Resume ImportFinalize

End Sub

Private Sub ProtectDashboardV2(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowSorting:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub
Public Sub RefreshTrendsV2()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Dim prevStatus As Variant
    prevStatus = Application.StatusBar

    Application.StatusBar = "方向フィルタを再計算しています..."

    On Error Resume Next
    UpdateAllDriverTrends ws
    ws.Calculate
    ApplyDynamicSignalsV2
    PreplaceOrdersV2
    If IsDemoMode() Then MarkPendingPreplaceAsOrderedDemo
    If IsDemoMode() Then ProcessDemoOrdersV2 ws
    UpdateTrendIndicators ws
    On Error GoTo 0

    Application.StatusBar = prevStatus

End Sub

Private Function IsDemoMode() As Boolean
    On Error Resume Next
    Dim nm As Name
    ' IMPORTANT: Use ThisWorkbook-scoped name first so behavior does not depend on
    ' which workbook is currently active/focused in Excel.
    Set nm = ThisWorkbook.Names("RunStatusV2")
    If nm Is Nothing Then Set nm = Application.Names("RunStatusV2")
    If Not nm Is Nothing Then
        IsDemoMode = (UCase$(Trim$(nm.RefersToRange.Value)) = "DEMO_RUNNING")
    Else
        IsDemoMode = False
    End If
    On Error GoTo 0
End Function

Private Sub MarkPendingPreplaceAsOrderedDemo()
    Dim Sh As Worksheet
    On Error Resume Next
    Set Sh = ThisWorkbook.Worksheets("Orders")
    On Error GoTo 0
    If Sh Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = Sh.Cells(Sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim modeVal As String: modeVal = UCase$(Trim$(CStr(Sh.Cells(r, 6).Value)))
        Dim statusVal As String: statusVal = UCase$(Trim$(CStr(Sh.Cells(r, 7).Value)))
        If modeVal = "PREPLACE" And statusVal = "PENDING" Then
            Sh.Cells(r, 6).Value = "preplace_demo"
            Sh.Cells(r, 7).Value = "ORDERED"
            If Len(Trim$(CStr(Sh.Cells(r, 8).Value))) = 0 Then
                Sh.Cells(r, 8).Value = "DEMO_PREPLACE"
            End If
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' DEMO: simulate fills and exits so Orders shows the full lifecycle.
' - preplace_demo ORDERED -> RUNNING when price is touched
' - when RUNNING, create tp_demo / sl_demo OCO orders (ORDERED)
' - when tp_demo/sl_demo is touched, mark exit FILLED and mark entry CLOSED
' ----------------------------------------------------------------------------

Private Sub ProcessDemoOrdersV2(ByVal ws As Worksheet)
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub
    If Not IsDemoMode() Then Exit Sub

    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("Orders")
    On Error GoTo Fail
    If sh Is Nothing Then Exit Sub

    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderTickerJP())
    If tickerCol = 0 Then Exit Sub
    Dim bestBidCol As Long: bestBidCol = FindColumn(ws, DASH2_HEADER_ROW, "BestBid")
    Dim bestAskCol As Long: bestAskCol = FindColumn(ws, DASH2_HEADER_ROW, "BestAsk")
    If bestBidCol = 0 Or bestAskCol = 0 Then Exit Sub

    ProcessDemoPreplaceFills ws, sh, tickerCol, bestBidCol, bestAskCol
    ProcessDemoExitFills ws, sh, tickerCol, bestBidCol, bestAskCol
    Exit Sub

Fail:
    LogVbaEvent "ProcessDemoOrdersV2", "Err " & Err.Number & ": " & Err.Description
End Sub

Private Function DashboardRowForTicker(ByVal ws As Worksheet, ByVal tickerCol As Long, ByVal ticker As String) As Long
    DashboardRowForTicker = 0
    If ws Is Nothing Or tickerCol = 0 Then Exit Function
    If Len(Trim$(ticker)) = 0 Then Exit Function
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, tickerCol).End(xlUp).Row
    Dim r As Long
    For r = DASH2_DATA_START To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, tickerCol).Value)), ticker, vbTextCompare) = 0 Then
            DashboardRowForTicker = r
            Exit Function
        End If
    Next r
End Function

Private Function HasOpenDemoPosition(ByVal sh As Worksheet, ByVal ticker As String) As Boolean
    HasOpenDemoPosition = False
    If sh Is Nothing Then Exit Function
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = lastRow To 2 Step -1
        If StrComp(Trim$(CStr(sh.Cells(r, ORD_COL_TICKER).Value)), ticker, vbTextCompare) = 0 Then
            Dim m As String: m = LCase$(Trim$(CStr(sh.Cells(r, ORD_COL_MODE).Value)))
            If m = "preplace_demo" Then
                Dim st As String: st = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_STATUS).Value)))
                If st = "RUNNING" Or st = "FILLED" Then
                    HasOpenDemoPosition = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Sub CancelOppositePreplaceDemo(ByVal sh As Worksheet, ByVal ticker As String, ByVal filledSide As String)
    If sh Is Nothing Then Exit Sub
    Dim cancelSide As String
    If StrComp(filledSide, "BUY", vbTextCompare) = 0 Then
        cancelSide = "SELL"
    ElseIf StrComp(filledSide, "SELL", vbTextCompare) = 0 Then
        cancelSide = "BUY"
    Else
        Exit Sub
    End If
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(sh.Cells(r, ORD_COL_TICKER).Value)), ticker, vbTextCompare) <> 0 Then GoTo NextCancel
        If StrComp(LCase$(Trim$(CStr(sh.Cells(r, ORD_COL_MODE).Value))), "preplace_demo", vbTextCompare) <> 0 Then GoTo NextCancel
        Dim st As String: st = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_STATUS).Value)))
        If st <> "ORDERED" And st <> "PENDING" Then GoTo NextCancel
        Dim side As String: side = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_SIDE).Value)))
        If side = UCase$(cancelSide) Then
            sh.Cells(r, ORD_COL_STATUS).Value = "CANCELLED_AUTO"
            If Len(Trim$(CStr(sh.Cells(r, ORD_COL_NOTE).Value))) = 0 Then
                sh.Cells(r, ORD_COL_NOTE).Value = "DEMO_OCO_CANCEL"
            End If
        End If
NextCancel:
    Next r
End Sub

Private Sub EnsureDemoExitOrders(ByVal sh As Worksheet, ByVal entryRow As Long)
    If sh Is Nothing Or entryRow <= 1 Then Exit Sub

    Dim ticker As String: ticker = Trim$(CStr(sh.Cells(entryRow, ORD_COL_TICKER).Value))
    Dim entrySide As String: entrySide = UCase$(Trim$(CStr(sh.Cells(entryRow, ORD_COL_SIDE).Value)))
    Dim qty As Double: qty = ToDouble(sh.Cells(entryRow, ORD_COL_QTY).Value, 0#)
    Dim tpPrice As Double: tpPrice = ToDouble(sh.Cells(entryRow, ORD_COL_TP).Value, 0#)
    Dim slPrice As Double: slPrice = ToDouble(sh.Cells(entryRow, ORD_COL_SL).Value, 0#)
    If Len(ticker) = 0 Or qty <= 0# Then Exit Sub

    Dim exitSide As String
    If entrySide = "BUY" Then
        exitSide = "SELL"
    ElseIf entrySide = "SELL" Then
        exitSide = "BUY"
    Else
        Exit Sub
    End If

    If tpPrice > 0# Then
        AppendDemoExitIfMissing sh, ticker, exitSide, tpPrice, qty, "tp_demo", "ORDERED", "DEMO_TP"
    End If
    If slPrice > 0# Then
        AppendDemoExitIfMissing sh, ticker, exitSide, slPrice, qty, "sl_demo", "ORDERED", "DEMO_SL"
    End If
End Sub

Private Sub AppendDemoExitIfMissing(ByVal sh As Worksheet, ByVal ticker As String, ByVal side As String, ByVal price As Double, ByVal qty As Double, ByVal mode As String, ByVal statusVal As String, ByVal note As String)
    If sh Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = lastRow To 2 Step -1
        If StrComp(Trim$(CStr(sh.Cells(r, ORD_COL_TICKER).Value)), ticker, vbTextCompare) = 0 Then
            If StrComp(UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_SIDE).Value))), UCase$(side), vbTextCompare) = 0 Then
                If StrComp(LCase$(Trim$(CStr(sh.Cells(r, ORD_COL_MODE).Value))), LCase$(mode), vbTextCompare) = 0 Then
                    Dim st As String: st = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_STATUS).Value)))
                    If st = "ORDERED" Or st = "PENDING" Or st = "FILLED" Or st = "RUNNING" Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next r

    Dim rowIdx As Long
    rowIdx = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1
    sh.Cells(rowIdx, ORD_COL_TS).Value = Now
    sh.Cells(rowIdx, ORD_COL_TICKER).Value = ticker
    sh.Cells(rowIdx, ORD_COL_SIDE).Value = UCase$(side)
    sh.Cells(rowIdx, ORD_COL_PRICE).Value = price
    sh.Cells(rowIdx, ORD_COL_QTY).Value = qty
    sh.Cells(rowIdx, ORD_COL_MODE).Value = mode
    sh.Cells(rowIdx, ORD_COL_STATUS).Value = statusVal
    sh.Cells(rowIdx, ORD_COL_NOTE).Value = note
    sh.Cells(rowIdx, ORD_COL_SOURCE).Value = "DEMO"
End Sub

Private Sub ProcessDemoPreplaceFills(ByVal ws As Worksheet, ByVal sh As Worksheet, ByVal tickerCol As Long, ByVal bestBidCol As Long, ByVal bestAskCol As Long)
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim modeVal As String: modeVal = LCase$(Trim$(CStr(sh.Cells(r, ORD_COL_MODE).Value)))
        If modeVal <> "preplace_demo" Then GoTo NextPreplace
        Dim statusVal As String: statusVal = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_STATUS).Value)))
        If statusVal <> "ORDERED" Then GoTo NextPreplace

        Dim ticker As String: ticker = Trim$(CStr(sh.Cells(r, ORD_COL_TICKER).Value))
        If Len(ticker) = 0 Then GoTo NextPreplace
        If HasOpenDemoPosition(sh, ticker) Then GoTo NextPreplace

        Dim dashRow As Long
        dashRow = DashboardRowForTicker(ws, tickerCol, ticker)
        If dashRow = 0 Then GoTo NextPreplace

        Dim bestBid As Double: bestBid = ToDouble(ws.Cells(dashRow, bestBidCol).Value, 0#)
        Dim bestAsk As Double: bestAsk = ToDouble(ws.Cells(dashRow, bestAskCol).Value, 0#)
        If bestBid <= 0# Or bestAsk <= 0# Then GoTo NextPreplace

        Dim side As String: side = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_SIDE).Value)))
        Dim limitPrice As Double: limitPrice = ToDouble(sh.Cells(r, ORD_COL_PRICE).Value, 0#)
        Dim qty As Double: qty = ToDouble(sh.Cells(r, ORD_COL_QTY).Value, 0#)
        If limitPrice <= 0# Or qty <= 0# Then GoTo NextPreplace

        Dim shouldFill As Boolean
        Dim fillPrice As Double
        If side = "BUY" Then
            shouldFill = (bestAsk <= limitPrice)
            If shouldFill Then fillPrice = IIf(bestAsk > 0#, bestAsk, limitPrice)
        ElseIf side = "SELL" Then
            shouldFill = (bestBid >= limitPrice)
            If shouldFill Then fillPrice = IIf(bestBid > 0#, bestBid, limitPrice)
        Else
            GoTo NextPreplace
        End If

        If Not shouldFill Then GoTo NextPreplace

        sh.Cells(r, ORD_COL_STATUS).Value = "RUNNING"
        sh.Cells(r, ORD_COL_FILL_TS).Value = Now
        sh.Cells(r, ORD_COL_FILL_PRICE).Value = fillPrice
        sh.Cells(r, ORD_COL_FILL_QTY).Value = qty
        sh.Cells(r, ORD_COL_SOURCE).Value = "DEMO"

        CancelOppositePreplaceDemo sh, ticker, side
        EnsureDemoExitOrders sh, r

NextPreplace:
    Next r
End Sub

Private Sub CancelSiblingExitOrders(ByVal sh As Worksheet, ByVal ticker As String, ByVal keepMode As String)
    If sh Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(sh.Cells(r, ORD_COL_TICKER).Value)), ticker, vbTextCompare) <> 0 Then GoTo NextExitCancel
        Dim modeVal As String: modeVal = LCase$(Trim$(CStr(sh.Cells(r, ORD_COL_MODE).Value)))
        If modeVal <> "tp_demo" And modeVal <> "sl_demo" Then GoTo NextExitCancel
        If modeVal = LCase$(keepMode) Then GoTo NextExitCancel
        Dim st As String: st = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_STATUS).Value)))
        If st = "ORDERED" Or st = "PENDING" Then
            sh.Cells(r, ORD_COL_STATUS).Value = "CANCELLED_AUTO"
            If Len(Trim$(CStr(sh.Cells(r, ORD_COL_NOTE).Value))) = 0 Then
                sh.Cells(r, ORD_COL_NOTE).Value = "DEMO_OCO_CANCEL"
            End If
        End If
NextExitCancel:
    Next r
End Sub

Private Sub ProcessDemoExitFills(ByVal ws As Worksheet, ByVal sh As Worksheet, ByVal tickerCol As Long, ByVal bestBidCol As Long, ByVal bestAskCol As Long)
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim modeVal As String: modeVal = LCase$(Trim$(CStr(sh.Cells(r, ORD_COL_MODE).Value)))
        If modeVal <> "tp_demo" And modeVal <> "sl_demo" Then GoTo NextExit
        Dim statusVal As String: statusVal = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_STATUS).Value)))
        If statusVal <> "ORDERED" Then GoTo NextExit

        Dim ticker As String: ticker = Trim$(CStr(sh.Cells(r, ORD_COL_TICKER).Value))
        If Len(ticker) = 0 Then GoTo NextExit

        Dim dashRow As Long
        dashRow = DashboardRowForTicker(ws, tickerCol, ticker)
        If dashRow = 0 Then GoTo NextExit

        Dim bestBid As Double: bestBid = ToDouble(ws.Cells(dashRow, bestBidCol).Value, 0#)
        Dim bestAsk As Double: bestAsk = ToDouble(ws.Cells(dashRow, bestAskCol).Value, 0#)
        If bestBid <= 0# Or bestAsk <= 0# Then GoTo NextExit

        Dim side As String: side = UCase$(Trim$(CStr(sh.Cells(r, ORD_COL_SIDE).Value)))
        Dim limitPrice As Double: limitPrice = ToDouble(sh.Cells(r, ORD_COL_PRICE).Value, 0#)
        Dim qty As Double: qty = ToDouble(sh.Cells(r, ORD_COL_QTY).Value, 0#)
        If limitPrice <= 0# Or qty <= 0# Then GoTo NextExit

        Dim shouldFill As Boolean
        Dim fillPrice As Double
        If modeVal = "tp_demo" Then
            If side = "SELL" Then
                shouldFill = (bestBid >= limitPrice)
                If shouldFill Then fillPrice = bestBid
            ElseIf side = "BUY" Then
                shouldFill = (bestAsk <= limitPrice)
                If shouldFill Then fillPrice = bestAsk
            End If
        ElseIf modeVal = "sl_demo" Then
            If side = "SELL" Then
                shouldFill = (bestBid <= limitPrice)
                If shouldFill Then fillPrice = bestBid
            ElseIf side = "BUY" Then
                shouldFill = (bestAsk >= limitPrice)
                If shouldFill Then fillPrice = bestAsk
            End If
        End If

        If Not shouldFill Then GoTo NextExit

        sh.Cells(r, ORD_COL_STATUS).Value = "FILLED"
        sh.Cells(r, ORD_COL_FILL_TS).Value = Now
        sh.Cells(r, ORD_COL_FILL_PRICE).Value = fillPrice
        sh.Cells(r, ORD_COL_FILL_QTY).Value = qty
        sh.Cells(r, ORD_COL_SOURCE).Value = "DEMO"

        Dim entrySide As String
        If side = "BUY" Then
            entrySide = "SELL"
        ElseIf side = "SELL" Then
            entrySide = "BUY"
        Else
            GoTo NextExit
        End If
        LogOrderSettled ticker, entrySide, fillPrice, qty, "DEMO_EXIT_" & UCase$(modeVal)
        CancelSiblingExitOrders sh, ticker, modeVal

NextExit:
    Next r
End Sub

Public Sub ClearBBBlocks()
    ResetBbBlockCache
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim entryStatusCol As Long: entryStatusCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryStatus")
    If entryStatusCol > 0 Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim r As Long
        For r = DASH2_DATA_START To lastRow
            Dim statusVal As String: statusVal = Trim$(CStr(ws.Cells(r, entryStatusCol).Value))
            If statusVal = "BLOCKED_BB" Or statusVal = "WARN_BB" Then
                ws.Cells(r, entryStatusCol).ClearContents
            End If
        Next r
    End If

    RefreshTrendsV2
End Sub


Private Sub EnsureParamFormulas(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    ' Build RSS labels without文字化け (use Unicode code points)
    Dim rssPriceNow As String
    rssPriceNow = ChrW(&H73FE) & ChrW(&H5728) & ChrW(&H5024)           ' 現在値
    Dim rssPctChange As String
    rssPctChange = ChrW(&H524D) & ChrW(&H65E5) & ChrW(&H6BD4) & ChrW(&H7387) ' 前日比率

    ' Index codes and RSS formulas (Japanese labels inside formula are fine)
    On Error Resume Next
    ws.Cells(2, 1).Value = "N225"
    ws.Cells(2, 4).Value = "TOPX"
    ws.Cells(2, 2).Formula = "=IF(A2="""","""",IFERROR(RssIndexMarket(A2,""" & rssPriceNow & """),""""))"
    ws.Cells(2, 3).Formula = "=IF(A2="""","""",IFERROR(RssIndexMarket(A2,""" & rssPctChange & """),""""))"
    ws.Cells(2, 5).Formula = "=IF(D2="""","""",IFERROR(RssIndexMarket(D2,""" & rssPriceNow & """),""""))"
    ws.Cells(2, 6).Formula = "=IF(D2="""","""",IFERROR(RssIndexMarket(D2,""" & rssPctChange & """),""""))"
    On Error GoTo 0

    ' Simple ASCII headers for parameter row (row 1)
    Dim paramHeaders As Variant
    paramHeaders = Array( _
        Array(1, "NKY_Code"), _
        Array(2, "NKY_Last"), _
        Array(3, "NKY_ChgPct"), _
        Array(4, "TOPIX_Code"), _
        Array(5, "TOPIX_Last"), _
        Array(6, "TOPIX_ChgPct"), _
        Array(7, "Bias_bp"), _
        Array(8, "BiasSlope"), _
        Array(9, "GapSlope"), _
        Array(10, "GapBanPct"), _
        Array(11, "NoTradeMin"), _
        Array(12, "TP_per_J"), _
        Array(13, "SL_per_J"), _
        Array(14, "Trail_per_J"), _
        Array(15, "CorrSlope"), _
        Array(16, "BudgetPerTicker"), _
        Array(17, "LotSize"), _
        Array(18, "NKY_TrendDay"), _
        Array(19, "NKY_TrendWindow"), _
        Array(20, "NKY_AllowedSide"), _
        Array(21, "TOPIX_TrendDay"), _
        Array(22, "TOPIX_TrendWindow"), _
        Array(23, "TOPIX_AllowedSide") _
    )

    Dim headerInfo As Variant
    For Each headerInfo In paramHeaders
        ws.Cells(1, headerInfo(0)).Value = headerInfo(1)
    Next headerInfo
End Sub
