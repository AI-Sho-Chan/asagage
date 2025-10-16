Attribute VB_Name = "FullAuto"
Option Explicit

'================= セクション：宣言ブロック（モジュールレベルに1回だけ） =================
' ---- OHLCパス ----
Public Const SIMPATH_STEPS      As Long = 600
Public Const SIMPATH_VWAP_ALPHA As Double = 0.15
Public Const SIMPATH_NOISE      As Double = 0.002
Public gSimPath As Object                                  ' code → state(Dictionary)

' ---- SIM 執行エンジンの設定値 ----
Public Const SIM_TH_ABSJ     As Double = 0.8
Public Const SIM_K_TP_ATR    As Double = 2#
Public Const SIM_K_SL_ATR    As Double = 2#
Public Const SIM_TMAX_BAR    As Long = 60
Public Const SIM_MIN_REENTRY As Long = 20
Public Const SIM_QTY         As Long = 100

' ---- 状態（起動フラグ／ポジション／ログ） ----
Public simTradeOn As Boolean
Public gSimPos    As Object                              ' Scripting.Dictionary
Public LiveBusy   As Boolean
Public CsvBusy    As Boolean
Private CompletedLogs As Object
Private Const LOG_MAX As Long = 1000

' ---- RSS 公式ラベル／学習済みラベル ----
Private Const PRICE_LABEL As String = "現在値"
Private Const VWAP_LABEL  As String = "出来高加重平均"
Private gPriceLabel As String
Private gVwapLabel  As String

' ---- 休場日シミュレータ状態 ----
Private simRunning As Boolean
Private simNext    As Date
Private simStepSec As Long
Private simRows    As Long
Private simVolPct  As Double
Private simSnap()  As Double

' ---- パス／シート ----
Public Const BASE_DIR    As String = "C:\AI\asagake"
Public Const SIGNAL_DIR  As String = BASE_DIR & "\data\signals\"
Public Const LIVE_DIR    As String = BASE_DIR & "\data\live"
Public Const LOG_DIR     As String = BASE_DIR & "\logs"
Public Const DASH_SHEET  As String = "Dashboard"
Public Const ORD_SHEET   As String = "Orders"
Public Const LOG_SHEET   As String = "LOG"
Public Const START_CELL  As String = "B5"

' ---- Live 周期/しきい ----
Public Const LIVE_INTERVAL_SEC As Long = 3
Public Const LIVE_J_MIN        As Double = 0.5

' ---- デモ/現物/信用 ----
Public DEMO_MODE   As Boolean
Public MARGIN_MODE As Boolean
Public Const LOT_DEFAULT As Long = 100

' ---- RSS注文 ----
Public Const SOR_KBN  As String = "0"
Public Const ACC_KOZA As String = "0"
Public g_OrderID      As Long
'================= セクション：宣言ブロック（ここまで） =================
'================= セクション：OnTimeスケジューラ（多段フォールバック） =================
' このブロックは宣言ブロックの直後に1回だけ配置
'================= セクション：OnTimeスケジューラ（最小版へルーティング） =================
Private Function ProcPatterns() As Variant
    Dim b$: b = ThisWorkbook.name
    ' ★ SimTick_Minimal を優先呼び出し
    ProcPatterns = Array( _
        "'" & b & "'!FullAuto.SimTick_Minimal", _
        "FullAuto.SimTick_Minimal")
End Function

Private Sub ReschedSim(Optional ByVal delaySec As Long = -1)
    If delaySec < 0 Then delaySec = simStepSec
    simNext = Now + TimeSerial(0, 0, delaySec)
    Dim p As Variant, ok As Boolean
    For Each p In ProcPatterns()
        On Error Resume Next
        Application.OnTime EarliestTime:=simNext, Procedure:=CStr(p), SCHEDULE:=True
        ok = (Err.Number = 0)
        On Error GoTo 0
        If ok Then
            LogEvent "SIMDBG", "ReschedOK", CStr(p) & " @" & Format$(simNext, "HH:NN:SS")
            Exit Sub
        End If
    Next
    LogEvent "SIMDBG", "ReschedFAIL", "all patterns failed"
End Sub

Private Sub CancelSim()
    If simNext = 0 Then Exit Sub
    Dim p As Variant
    For Each p In ProcPatterns()
        On Error Resume Next
        Application.OnTime EarliestTime:=simNext, Procedure:=CStr(p), SCHEDULE:=False
        On Error GoTo 0
    Next
    simNext = 0
End Sub
'================= セクション終わり =================



'==========================================================
' SHINSOKU ? 安定版フルモジュール
' ・Signalsの自動解決（今日が無ければ最新日）
' ・CSV取り込み（UTF-8/ダブルクォート対応の堅牢パーサ）
' ・Ticker列名の揺れ吸収（Ticker/code/銘柄コード/Code）
' ・RssMarket：公式ラベル「現在値」「出来高加重平均」を最優先→保険ラベル総当たり
' ・Live(3秒)で J/dJ/vEMA を必ず更新（値が取れた行のみ）
' ・ワンクリック検証 LiveOnce_Debug
' ・デモ発注：CSV保存＋Ordersに即追記
' ・BookNew：完全リセット＋ボタン再割当
' ・UIハンドラ（重複名を避け一意に）
' ・日次PS（Pipeline）起動用：BtnRunPipeline_Click + PSユーティリティ
'==========================================================

'================= セクション：SmoothStep（置換） =================
Private Function SmoothStep(ByVal u As Double) As Double
    If u < 0# Then
        u = 0#
    ElseIf u > 1# Then
        u = 1#
    End If
    SmoothStep = u * u * (3# - 2# * u)
End Function
'================= セクション：SmoothStep（ここまで） =================




' OHLC に沿って「Open→(High/Low)→(Low/High)→Close」を通過する滑らかパス
' ・codeごとに状態を保持（gSimPath: Dictionary）
' ・戻り値：次の Close 値、vwap は ByRef で返す
'================= セクション：SimPath_NextPrice_OHLC（擬似HL帯＋Seg進行） =================
Private Function SimPath_NextPrice_OHLC(ByVal ws As Worksheet, ByVal r As Long, _
    ByVal cOpen As Long, ByVal cHigh As Long, ByVal cLow As Long, ByVal cClose As Long, _
    ByVal cVW As Long, ByVal code As String, ByRef vwap As Double) As Double

    If gSimPath Is Nothing Then Set gSimPath = CreateObject("Scripting.Dictionary")

    Dim o#, h#, l#, c#
    o = d(IIf(cOpen > 0, ws.Cells(r, cOpen).value, ws.Cells(r, cClose).value))
    h = d(IIf(cHigh > 0, ws.Cells(r, cHigh).value, o))
    l = d(IIf(cLow > 0, ws.Cells(r, cLow).value, o))
    c = d(ws.Cells(r, cClose).value): If c <= 0 Then c = o

    ' HLが無効なら擬似レンジを生成（±max(0.05%, simVolPct)）
    If h <= l Then
        Dim base#, band#
        base = IIf(Abs(o) > 0, Abs(o), IIf(Abs(c) > 0, Abs(c), 100#))
        band = Application.Max(base * 0.0005, base * simVolPct)
        h = base + band: l = base - band
        If o <= 0 Then o = base
        If c <= 0 Then c = base
    End If

    Dim state As Object
    If Not gSimPath.Exists(code) Then
        Set state = CreateObject("Scripting.Dictionary")
        state("v1") = o
        state("v2") = IIf(Rnd < 0.5, h, l)
        state("v3") = IIf(state("v2") = h, l, h)
        state("v4") = c
        state("t1") = Int(SIMPATH_STEPS * (0.2 + 0.4 * Rnd))
        state("t2") = Int(SIMPATH_STEPS * (0.6 + 0.2 * Rnd))
        If state("t2") <= state("t1") Then state("t2") = state("t1") + 1
        state("t") = 0: state("seg") = 1
        vwap = d(IIf(cVW > 0, ws.Cells(r, cVW).value, o))
        state("vwap") = vwap
        gSimPath(code) = state
    Else
        Set state = gSimPath(code)
        vwap = state("vwap")
    End If

    Dim t&, seg&, t1&, t2&, p0#, p1#, p2#, p3#
    t = state("t"): seg = state("seg")
    t1 = state("t1"): t2 = state("t2")
    p0 = state("v1"): p1 = state("v2"): p2 = state("v3"): p3 = state("v4")

    ' 1ティック進める
    t = t + 1: If t > SIMPATH_STEPS Then t = SIMPATH_STEPS

    Dim t0&, tA&, u#, price#
    Select Case seg
        Case 1: t0 = 0:   tA = t1
        Case 2: t0 = t1:  tA = t2
        Case Else: t0 = t2: tA = SIMPATH_STEPS
    End Select

    u = (t - t0) / (tA - t0)
    If u < 0# Then
        u = 0#
    ElseIf u > 1# Then
        u = 1#
    End If

    price = Choose(seg, _
        p0 + (p1 - p0) * SmoothStep(u), _
        p1 + (p2 - p1) * SmoothStep(u), _
        p2 + (p3 - p2) * SmoothStep(u))

    ' 次セグメントへ進行
    If t >= tA Then
        If seg < 3 Then seg = seg + 1
    End If

    ' 微小ノイズ＋クランプ
    price = price * (1 + RandN() * SIMPATH_NOISE)
    If price < l Then price = l
    If price > h Then price = h

    ' VWAP 追随
    vwap = vwap + SIMPATH_VWAP_ALPHA * (price - vwap)

    ' 保存
    state("t") = t: state("seg") = seg: state("vwap") = vwap
    gSimPath(code) = state

    SimPath_NextPrice_OHLC = price
End Function
'================= セクション終わり：SimPath_NextPrice_OHLC =================




'=== UI handlers（ボタン用。名前衝突しないよう UI_ 接頭辞に統一） ===
Public Sub UI_Import_Click():             ImportTopCSV: End Sub
Public Sub UI_AutoStartDemo_Click():      AutoStart True: End Sub
Public Sub UI_AutoStartLive_Click():      AutoStart False: End Sub
Public Sub UI_AutoStop_Click():           AutoStop: End Sub
Public Sub UI_RefreshOrdersPanel_Click(): RefreshOrdersPanel: End Sub
'（日次PSボタンを使うなら）
'Public Sub UI_RunPipeline_Click():        BtnRunPipeline_Click: End Sub


'----------------------------------------------------------
' 汎用ユーティリティ
'----------------------------------------------------------
Private Sub EnsureLogInfra()
    If CompletedLogs Is Nothing Then Set CompletedLogs = CreateObject("Scripting.Dictionary")
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function

Private Sub EnsureDirs()
    On Error Resume Next
    If Dir$(LOG_DIR, vbDirectory) = "" Then MkDir LOG_DIR
    If Dir$(BASE_DIR & "\data", vbDirectory) = "" Then MkDir (BASE_DIR & "\data")
    If Dir$(LIVE_DIR, vbDirectory) = "" Then MkDir LIVE_DIR
End Sub

Private Sub LogEvent(ByVal kind As String, ByVal detail As String, Optional ByVal note As String = "")
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(LOG_SHEET): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add: ws.name = LOG_SHEET
        ws.Range("A1:D1").value = Array("Time", "Kind", "Detail", "Note")
        ws.Range("A1:D1").Font.Bold = True
        ws.Columns("A:D").ColumnWidth = 26
    End If
    Dim last As Long: last = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If last >= LOG_MAX Then ws.rows(2).Delete
    ws.Cells(last + 1, 1).Resize(1, 4).value = _
        Array(Format$(Now, "yyyy-MM-dd HH:mm:ss"), kind, detail, note)
End Sub

Private Function SetStatus(ByVal s As String) As Boolean
    On Error Resume Next
    With ThisWorkbook.Worksheets(DASH_SHEET)
        .Range("A3").value = "STATUS: " & s & "  @ " & Format$(Now, "hh:nn:ss")
    End With
    Application.StatusBar = s
    SetStatus = True
End Function

Private Function FindCol(h As Range, title As String) As Long
    Dim c As Range
    For Each c In h.Cells
        If CStr(c.value) = title Then FindCol = c.Column: Exit Function
    Next c
    FindCol = 0
End Function

' 見出し候補のうち最初に見つかった列番号を返す（見つかったタイトル名も返す）
Private Function FindColAny(h As Range, ByRef foundTitle As String, ParamArray titles()) As Long
    Dim t As Variant, c As Range
    foundTitle = ""
    For Each t In titles
        For Each c In h.Cells
            If StrComp(CStr(c.value), CStr(t), vbTextCompare) = 0 Then
                FindColAny = c.Column
                foundTitle = CStr(t)
                Exit Function
            End If
        Next c
    Next t
    FindColAny = 0
End Function

Private Function d(ByVal v As Variant) As Double
    If IsNumeric(v) Then d = CDbl(v) Else d = Val(Replace(CStr(v), ",", ""))
End Function

Private Function TickerToCode(ByVal tkr As String) As String
    Dim i As Long: i = InStr(1, tkr, ".", vbTextCompare)
    If i > 0 Then TickerToCode = Left$(tkr, i - 1) Else TickerToCode = tkr
End Function

'----------------------------------------------------------
' Signalsパス：今日が無ければ最新日
'----------------------------------------------------------
Private Function ResolveSignalsPath() As String
    Dim today As String, p As String
    today = Format(Date, "yyyy-mm-dd")
    p = SIGNAL_DIR & today & "\top_candidates.csv"
    If Dir$(p) <> "" Then ResolveSignalsPath = p: Exit Function

    Dim best As String, f As String, d As Date, bestD As Date
    f = Dir$(SIGNAL_DIR, vbDirectory)
    Do While Len(f) > 0
        If (GetAttr(SIGNAL_DIR & f) And vbDirectory) = vbDirectory Then
            If f <> "." And f <> ".." Then
                If IsDate(f) Then
                    If Dir$(SIGNAL_DIR & f & "\top_candidates.csv") <> "" Then
                        d = CDate(f)
                        If d > bestD Then bestD = d: best = SIGNAL_DIR & f & "\top_candidates.csv"
                    End If
                End If
            End If
        End If
        f = Dir$
    Loop
    ResolveSignalsPath = best
End Function
'----------------------------------------------------------
' CSV 取込（UTF-8/ダブルクォート対応）＋ 数式敷設
'----------------------------------------------------------
Public Sub ImportTopCSV()
    On Error GoTo FAIL

    EnsureDirs
    Dim p As String: p = ResolveSignalsPath()
    If Len(p) = 0 Or Dir$(p) = "" Then
        MsgBox "CSVが見つかりません: " & SIGNAL_DIR, vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    ws.Range(START_CELL).CurrentRegion.ClearContents
    ws.Range("A2").value = "Signals: " & p

    ' --- UTF-8 読み込み ---
    Dim stm As Object, txt As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile p
    txt = stm.ReadText(-1)
    stm.Close: Set stm = Nothing

    ' --- CSV → 2次元配列（ダブルクォート対応）---
    Dim raw() As String, lines As Object, i As Long
    Dim hdr() As String, cols As Long, rCount As Long
    Dim arr() As Variant, vals As Variant, j As Long, ub As Long

    raw = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    Set lines = CreateObject("System.Collections.ArrayList")
    For i = 0 To UBound(raw)
        If Len(Trim$(raw(i))) > 0 Then lines.Add raw(i)
    Next i
    If lines.Count < 2 Then GoTo FAIL

    hdr = ParseCsvLine(CStr(lines(0)))
    cols = UBound(hdr) + 1
    rCount = lines.Count - 1
    ReDim arr(1 To rCount, 1 To cols)

    Dim lineText As String
    For i = 1 To lines.Count - 1
        lineText = CStr(lines(i))
        vals = ParseCsvLine(lineText)
        On Error Resume Next: ub = UBound(vals): On Error GoTo 0
        For j = 1 To cols
            If ub >= 0 And j <= ub + 1 Then
                arr(i, j) = TryNum(vals(j - 1))
            Else
                arr(i, j) = ""
            End If
        Next j
    Next i

    ' --- シートへ貼付 ---
    ws.Range(START_CELL).Resize(1, cols).value = hdr
    ws.Range(START_CELL).Offset(1, 0).Resize(rCount, cols).value = arr

    ' ts列（B列）を文字列固定
    ws.Columns(ws.Range(START_CELL).Column + 1).NumberFormatLocal = "@"

    ' コード列の告知（Ticker/code/銘柄コード/Code のいずれか）
    Dim hdrRow As Range: Set hdrRow = ws.Range(START_CELL).Resize(1, cols)
    Dim reportCol As String, cT As Long
    cT = FindColAny(hdrRow, reportCol, "Ticker", "code", "銘柄コード", "Code")
    If cT = 0 Then
        LogEvent "CSV", "NoCodeColumn", "top_candidates.csv にコード列がありません（Ticker/code/銘柄コード/Code）"
    Else
        LogEvent "CSV", "CodeColumn", "Use=" & reportCol
    End If
'================= セクション末尾：ImportTopCSV の終端（置換） =================
    ' ★ J/dJ/vEMA の数式を敷設（CSV取込の度に最新行まで張り直し）
    InstallJKVFormulas False

    ' ★ 取り込み直後にOHLC補修（桁ズレ・HL整合）
    SanitizeOHLC_Repair

    LogEvent "CSV", "ImportTopCSV", p
    SetStatus "CSV取込完了"
    Exit Sub
' 以降 FAIL: は現行どおり
'================= セクション終わり：ImportTopCSV 終端 =================

FAIL:
    LogEvent "ERR", "ImportTopCSV", Err.Description
    SetStatus "CSV取込エラー"
End Sub


Private Function ParseCsvLine(ByVal s As String) As String()
    Dim out As Object: Set out = CreateObject("System.Collections.ArrayList")
    Dim i As Long, cH As String, cur As String, inQ As Boolean
    inQ = False: cur = ""
    For i = 1 To Len(s)
        cH = Mid$(s, i, 1)
        If cH = """" Then
            If inQ And i < Len(s) And Mid$(s, i + 1, 1) = """" Then
                cur = cur & """": i = i + 1
            Else
                inQ = Not inQ
            End If
        ElseIf cH = "," And Not inQ Then
            out.Add cur: cur = ""
        Else
            cur = cur & cH
        End If
    Next i
    out.Add cur
    Dim arr() As String, n As Long, k As Long
    n = out.Count: ReDim arr(0 To n - 1)
    For k = 0 To n - 1: arr(k) = out(k): Next
    ParseCsvLine = arr
End Function

Private Function TryNum(ByVal x As String) As Variant
    Dim t As String: t = Trim$(x)
    If t = "" Then TryNum = "": Exit Function
    t = Replace(t, ",", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, ChrW(&H3000), "")
    If IsNumeric(t) Then TryNum = CDbl(t) Else TryNum = x
End Function

'================= セクション：SanitizeOHLC_Repair（CSVのOHLC補修） =================
' 目的：
'   ・High < max(Open,Close) / Low > min(Open,Close) を補修
'   ・桁ズレ（×10, ×100）の自動補正（VWAP/Close との比で判定）
'   ・VWAP欠損時は Close で代用
' 呼び出し：
'   ImportTopCSV の最後で SanitizeOHLC_Repair を呼ぶ
'================= セクション：SanitizeOHLC_Repair（名前衝突解消版） =================

'================= セクション：SanitizeOHLC_Repair（衝突回避・ASCII） =================
Private Sub SanitizeOHLC_Repair()
    Dim ws As Worksheet, rg As Range, hd As Range
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Set rg = ws.Range(START_CELL).CurrentRegion
    If rg Is Nothing Or rg.rows.Count < 2 Then Exit Sub
    Set hd = rg.rows(1)

    ' 列インデックス（col*）
    Dim colO As Long, colH As Long, colL As Long, colC As Long, colVW As Long
    colO = FindCol(hd, "Open")
    colH = FindCol(hd, "High")
    colL = FindCol(hd, "Low")
    colC = FindCol(hd, "Close")
    colVW = FindCol(hd, "VWAPd"): If colVW = 0 Then colVW = FindCol(hd, "VWAP")
    If colC = 0 Then Exit Sub

    ' 行範囲
    Dim r1 As Long, r2 As Long, r As Long
    r1 = rg.row + 1
    r2 = ws.Cells(ws.rows.Count, colC).End(xlUp).row
    If r2 < r1 Then Exit Sub

    ' 値（v*）
    Dim vO As Double, vH As Double, vL As Double, vC As Double, vVW As Double
    Dim mx As Double, mn As Double, rat As Double, sc As Double, t As Double

    For r = r1 To r2
        vO = IIf(colO > 0, d(ws.Cells(r, colO).value), 0#)
        vH = IIf(colH > 0, d(ws.Cells(r, colH).value), 0#)
        vL = IIf(colL > 0, d(ws.Cells(r, colL).value), 0#)
        vC = d(ws.Cells(r, colC).value)
        vVW = IIf(colVW > 0, d(ws.Cells(r, colVW).value), 0#)
        If vVW <= 0# Then vVW = vC

        ' 1) 桁ズレ補正（×10/×100）
        If vC > 0# And vVW > 0# Then
            rat = vC / vVW
            If rat > 5# Then
                sc = 10#
                If vC / sc / vVW > 0.2 And vC / sc / vVW < 5# Then
                    vC = vC / sc: vO = vO / sc: vH = vH / sc: vL = vL / sc
                ElseIf vC / (sc * sc) / vVW > 0.2 And vC / (sc * sc) / vVW < 5# Then
                    vC = vC / (sc * sc)
                    vO = vO / (sc * sc): vH = vH / (sc * sc): vL = vL / (sc * sc)
                End If
            ElseIf rat < 0.2 Then
                sc = 10#
                If vC * sc / vVW > 0.2 And vC * sc / vVW < 5# Then
                    vC = vC * sc: vO = vO * sc: vH = vH * sc: vL = vL * sc
                ElseIf vC * (sc * sc) / vVW > 0.2 And vC * (sc * sc) / vVW < 5# Then
                    vC = vC * (sc * sc)
                    vO = vO * (sc * sc): vH = vH * (sc * sc): vL = vL * (sc * sc)
                End If
            End If
        End If

        ' 2) HL整合
        If colO = 0 Then vO = vC
        If colH = 0 Then vH = vC
        If colL = 0 Then vL = vC
        mx = WorksheetFunction.Max(vO, vC)
        mn = WorksheetFunction.Min(vO, vC)
        If vH < mx Then vH = mx * 1.0001
        If vL > mn Then vL = mn * 0.9999
        If vH < vL Then t = vH: vH = vL: vL = t

        ' 3) 書戻し
        If colO > 0 Then ws.Cells(r, colO).Value2 = vO
        If colH > 0 Then ws.Cells(r, colH).Value2 = vH
        If colL > 0 Then ws.Cells(r, colL).Value2 = vL
        ws.Cells(r, colC).Value2 = vC
        If colVW > 0 And vVW > 0 Then ws.Cells(r, colVW).Value2 = vVW
    Next r

    LogEvent "CSV", "SanitizeOHLC", "rows=" & (r2 - r1 + 1)
End Sub
'================= セクション終わり：SanitizeOHLC_Repair =================


'----------------------------------------------------------
' RssMarket 取得（学習済み→公式→保険）
'----------------------------------------------------------
Private Function RSS_Get(ByVal code As String, ByVal want As String, Optional ByRef outName As String = "") As Double
    Dim fnames As Variant, fields As Variant
    Dim i As Long, j As Long, v As Variant
    fnames = Split("RssMarket,RssMarket_v,RssMarketEx", ",")

    ' 学習済み優先
    If want = "PRICE_CAND" And Len(gPriceLabel) > 0 Then
        For i = LBound(fnames) To UBound(fnames)
            On Error Resume Next: v = Application.Run(CStr(fnames(i)), code, gPriceLabel): On Error GoTo 0
            If IsNumeric(v) And CDbl(v) > 0 Then outName = gPriceLabel: RSS_Get = CDbl(v): Exit Function
        Next i
    ElseIf want = "VWAP_CAND" And Len(gVwapLabel) > 0 Then
        For i = LBound(fnames) To UBound(fnames)
            On Error Resume Next: v = Application.Run(CStr(fnames(i)), code, gVwapLabel): On Error GoTo 0
            If IsNumeric(v) And CDbl(v) > 0 Then outName = gVwapLabel: RSS_Get = CDbl(v): Exit Function
        Next i
    End If

    ' 公式ラベル
    Dim target As String
    If want = "PRICE_CAND" Then
        target = PRICE_LABEL
    ElseIf want = "VWAP_CAND" Then
        target = VWAP_LABEL
    Else
        target = want
    End If
    For i = LBound(fnames) To UBound(fnames)
        On Error Resume Next: v = Application.Run(CStr(fnames(i)), code, target): On Error GoTo 0
        If IsNumeric(v) And CDbl(v) > 0 Then outName = target: RSS_Get = CDbl(v): Exit Function
    Next i

    ' 3) 保険ラベル（同義表記を総当たり）
Dim priceCands As String, vwapCands As String
priceCands = PRICE_LABEL & ",現在値(気配),現値,直近値,Last,LastPrice"
vwapCands = VWAP_LABEL & ",出来高加重平均(価格),出来高加重平均（価格）,VWAP,加重平均,加重平均(価格)"

If want = "PRICE_CAND" Then
    fields = Split(priceCands, ",")
Else
    fields = Split(vwapCands, ",")
End If

For i = LBound(fnames) To UBound(fnames)
    For j = LBound(fields) To UBound(fields)
        On Error Resume Next
        v = Application.Run(CStr(fnames(i)), code, CStr(fields(j)))
        On Error GoTo 0

        If IsNumeric(v) And CDbl(v) > 0 Then
            outName = CStr(fields(j))
            RSS_Get = CDbl(v)

            ' 学習保存（ブロックIfで安全に）
            If want = "PRICE_CAND" Then
                gPriceLabel = outName
            ElseIf want = "VWAP_CAND" Then
                gVwapLabel = outName
            End If

            Exit Function
        End If
    Next j
Next i

    outName = "": RSS_Get = 0#
End Function

' 起動時：先頭銘柄でラベル自動解決
Public Sub ResolveRssLabels(Optional ByVal sampleCode As String = "")
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:   Set rng = ws.Range(START_CELL).CurrentRegion
    If Len(sampleCode) = 0 And rng.rows.Count >= 2 Then
        Dim tit As String, cT As Long
        cT = FindColAny(rng.rows(1), tit, "Ticker", "code", "銘柄コード", "Code")
        If cT > 0 Then
            sampleCode = TickerToCode(CStr(rng.Cells(2, cT).value))
        Else
            LogEvent "RESOLVE", "NoCodeColumn", "Header not found (Ticker/code/銘柄コード/Code)"
            Exit Sub
        End If
    End If
    If Len(sampleCode) <> 4 Then LogEvent "RESOLVE", "Skip", "No sample code": Exit Sub

    Dim n1 As String, n2 As String, px As Double, vw As Double
    px = RSS_Get(sampleCode, "PRICE_CAND", n1)
    vw = RSS_Get(sampleCode, "VWAP_CAND", n2)
    If px > 0 Then gPriceLabel = n1
    If vw > 0 Then gVwapLabel = n2
    LogEvent "RESOLVE", "Labels", "Price=" & IIf(px > 0, n1, "(none)") & " VWAP=" & IIf(vw > 0, n2, "(none)")
End Sub

'----------------------------------------------------------
' Live（3秒）：J/dJ/vEMA を更新
'----------------------------------------------------------
Public Sub LiveStart()
    If LiveBusy Then Exit Sub
    LogEvent "SYS", "LiveStart", ""
    LiveTick
End Sub

Public Sub LiveStop()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC), "LiveTick", , False
    LogEvent "SYS", "LiveStop", ""
End Sub

Public Sub LiveTick()
    Static beat As Long
    If CsvBusy Then GoTo RESCHEDULE
    LiveBusy = True

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng.rows.Count >= 2 Then
        Dim h As Range: Set h = rng.rows(1)
        Dim cT&, cC&, cA&, cJ&, cdJ&, cvEMA&, cVW&, titCode$
        ' 列検出（Ticker 名の揺れ対応）
        cT = FindColAny(h, titCode, "Ticker", "code", "銘柄コード", "Code")
        cC = FindCol(h, "Close"): cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J")
        cdJ = FindCol(h, "dJ"):  cvEMA = FindCol(h, "vEMA")
        cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

        If cT * cC * cA * cJ = 0 Then
            LogEvent "LIVEDBG", "HeaderMissing", _
                     "RowCnt=" & rng.rows.Count & " cT=" & cT & "(" & titCode & ")" & _
                     " cC=" & cC & " cA=" & cA & " cJ=" & cJ
            GoTo RESCHEDULE
        End If

        Dim r As Long
        For r = 2 To rng.rows.Count
            Dim tkr$, code$, px#, vw#, atr#, jp#, jn#, dj#, ve#, nPx$, nVw$
            tkr = Trim$(CStr(rng.Cells(r, cT).value)): If Len(tkr) = 0 Then GoTo NXT
            code = TickerToCode(tkr):                 If Len(code) <> 4 Then GoTo NXT

            ' 値の取得（公式ラベル優先→候補フォールバック）
            px = RSS_Get(code, "PRICE_CAND", nPx)
            vw = RSS_Get(code, "VWAP_CAND", nVw)
            atr = d(rng.Cells(r, cA).value)
            If px <= 0# Or vw <= 0# Or atr <= 0# Then GoTo NXT

            ' 旧J
            jp = d(rng.Cells(r, cJ).value)

            ' === 書込みセクション（ここから） ===
            jn = (px - vw) / atr
            dj = jn - jp
            ve = 0.3 * dj + 0.7 * d(IIf(cvEMA > 0, rng.Cells(r, cvEMA).value, 0))
' --- LiveTick 書込み部（ここから置換） ---
' Close / VWAPd は値で更新
rng.Cells(r, cC).Value2 = px
If cVW > 0 Then rng.Cells(r, cVW).Value2 = vw

' J列が数式なら、一旦値化してから上書き（数式セルでも必ず入る）
If rng.Cells(r, cJ).HasFormula Then rng.Cells(r, cJ).Value2 = rng.Cells(r, cJ).Value2
' J → dJ → vEMA の順で強制上書き（Value2）
rng.Cells(r, cJ).Value2 = jn
If cdJ > 0 Then rng.Cells(r, cdJ).Value2 = dj
If cvEMA > 0 Then rng.Cells(r, cvEMA).Value2 = ve
' --- LiveTick 書込み部（置換ここまで） ---


NXT:    Next r
    End If

RESCHEDULE:
    LiveBusy = False
    Application.OnTime Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC), "LiveTick"
End Sub


' ワンクリック検証（寄り前でもJ/dJ/vEMAの更新を確認）
Public Sub LiveOnce_Debug()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:   Set rng = ws.Range(START_CELL).CurrentRegion
    If rng.rows.Count < 2 Then MsgBox "候補がありません": Exit Sub

    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cvEMA&, cVW&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "銘柄コード", "Code")
    cC = FindCol(h, "Close"): cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J")
    cdJ = FindCol(h, "dJ"): cvEMA = FindCol(h, "vEMA"): cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    If cT * cC * cA * cJ = 0 Then
        LogEvent "LIVEDBG", "HeaderMissing", "cT=" & cT & "(" & tit & ") cC=" & cC & " cA=" & cA & " cJ=" & cJ
        Exit Sub
    End If

    Dim r As Long, cnt As Long, tkr$, code$, px#, vw#, atr#, jp#, jn#, dj#, ve#, nPx$, nVw$, rep$
    For r = 2 To Application.Min(rng.rows.Count, 6)
        tkr = Trim$(CStr(rng.Cells(r, cT).value)): If Len(tkr) = 0 Then GoTo NXT
        code = TickerToCode(tkr):                 If Len(code) <> 4 Then GoTo NXT

        px = RSS_Get(code, "PRICE_CAND", nPx)
        vw = RSS_Get(code, "VWAP_CAND", nVw)
        atr = d(rng.Cells(r, cA).value)

        If px > 0 And vw > 0 And atr > 0 Then
            jp = d(rng.Cells(r, cJ).value)
            jn = (px - vw) / atr: dj = jn - jp
            ve = 0.3 * dj + 0.7 * d(IIf(cvEMA > 0, rng.Cells(r, cvEMA).value, 0))

            rng.Cells(r, cC).value = px
            rng.Cells(r, cJ).value = jn
            If cdJ > 0 Then rng.Cells(r, cdJ).value = dj
            If cvEMA > 0 Then rng.Cells(r, cvEMA).value = ve
            If cVW > 0 Then rng.Cells(r, cVW).value = vw

            rep = rep & IIf(Len(rep) > 0, " | ", "") & code & _
                  ":px=" & Format$(px, "0.###") & "(" & nPx & ")" & _
                  " vw=" & Format$(vw, "0.###") & "(" & nVw & ")" & _
                  " J:" & Format$(jp, "0.###") & "→" & Format$(jn, "0.###")
            cnt = cnt + 1
        Else
            rep = rep & IIf(Len(rep) > 0, " | ", "") & code & ":px/vw/atr=0"
        End If
NXT:
    Next r
    LogEvent "LIVEDBG", "LiveOnce", rep
    MsgBox "LiveOnce_Debug 完了：LOG の LIVEDBG | LiveOnce を確認してください。", vbInformation
End Sub

'----------------------------------------------------------
' デモ発注：CSV保存＋Ordersに即追記
'----------------------------------------------------------
Private Sub WriteDemoLog(ByVal code As String, ByVal side As String, ByVal qty As Long, ByVal price As Double, ByVal p As Double, Optional ByVal status As String = "")
    EnsureDirs
    Dim fn As String: fn = LOG_DIR & "\demo_orders.csv": Dim f As Integer: f = FreeFile
    If Dir$(fn) = "" Then Open fn For Output As #f: Print #f, "ts,mode,margin,code,side,qty,price,p,status" _
                   Else Open fn For Append As #f
    Print #f, Format$(Now, "yyyy-MM-dd HH:mm:ss") & "," & IIf(DEMO_MODE, "DEMO", "LIVE") & "," & IIf(MARGIN_MODE, "MARGIN", "CASH") & "," & _
             code & "," & side & "," & qty & "," & Format$(price, "0.000") & "," & Format$(p, "0.000") & "," & status
    Close #f

    Dim ws As Worksheet: On Error Resume Next: Set ws = ThisWorkbook.Worksheets(ORD_SHEET): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Trim$(ws.Range("A1").value) <> "■ 注文一覧" Then
        ws.Cells.Clear
        ws.Range("A1").value = "■ 注文一覧": ws.Range("A1").Font.Bold = True
        ws.Range("A4").value = "■ 約定一覧": ws.Range("A4").Font.Bold = True
        ws.Range("A2").value = "(デモ注文ログがありません)"
    End If
    Dim row As Long: row = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If row < 2 Then row = 1
    If ws.Range("A2").value = "(デモ注文ログがありません)" Then ws.Range("A2").ClearContents: row = 1
    With ws
        .Cells(row + 1, 1).value = Format$(Now, "yyyy-MM-dd HH:mm:ss")
        .Cells(row + 1, 2).value = "DEMO-ORDER"
        .Cells(row + 1, 3).value = side & " " & code & " x" & qty & " @" & Format$(price, "0.000")
        .Cells(row + 1, 4).value = status
    End With
End Sub

'----------------------------------------------------------
' Ordersパネル（必要に応じて本番のRSS一覧に差替可）
'----------------------------------------------------------
Public Sub RefreshOrdersPanel()
    Dim ws As Worksheet: On Error Resume Next: Set ws = ThisWorkbook.Worksheets(ORD_SHEET): On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.name = ORD_SHEET
    ws.Columns.AutoFit
End Sub

'----------------------------------------------------------
' BookNew：フルリセット＋ボタン再割当＋J/dJ/vEMA の数式敷設
'----------------------------------------------------------
Public Sub BookNew()
    On Error Resume Next

    '---- 1) 安全停止（タイマー解除）
    Application.OnTime Now + TimeSerial(0, 0, 1), "LiveTick", , False
    LiveBusy = False: CsvBusy = False

    '---- 2) フォルダ確保
    EnsureDirs

    '---- 3) Dashboard を再生成
    Dim wsD As Worksheet
    If SheetExists(DASH_SHEET) Then
        Set wsD = ThisWorkbook.Worksheets(DASH_SHEET)
        wsD.Cells.Clear
    Else
        Set wsD = ThisWorkbook.Worksheets.Add
        wsD.name = DASH_SHEET
    End If

    With wsD
        .Range("A1").value = "逆張りスキャルピング Dashboard（BookNew）"
        .Range("A2").value = "Signals: " & ResolveSignalsPath()
       .Range(START_CELL).Resize(1, 24).value = Array( _
    "ts", "Open", "High", "Low", "Close", "Volume", "VWAPd", "ATR5", "J", "dJ", "d2J", "vEMA", _
    "Session", "Ticker", "IBS", "SMA20", "STD20", "Z20", "ROC5", "Turnover", "p", "side", "TP_price", "SL_price", "Value")
        .Columns.ColumnWidth = 12
    End With

    '---- 4) LOG シート
    Dim wsL As Worksheet
    If SheetExists(LOG_SHEET) Then
        Set wsL = ThisWorkbook.Worksheets(LOG_SHEET): wsL.Cells.Clear
    Else
        Set wsL = ThisWorkbook.Worksheets.Add: wsL.name = LOG_SHEET
    End If
    wsL.Range("A1:D1").value = Array("Time", "Kind", "Detail", "Note")
    wsL.Range("A1:D1").Font.Bold = True
    wsL.Columns("A:D").ColumnWidth = 26

    '---- 5) Orders シート
    Dim wsO As Worksheet
    If SheetExists(ORD_SHEET) Then
        Set wsO = ThisWorkbook.Worksheets(ORD_SHEET): wsO.Cells.Clear
    Else
        Set wsO = ThisWorkbook.Worksheets.Add: wsO.name = ORD_SHEET
    End If
    With wsO
        .Range("A1").value = "■ 注文一覧": .Range("A1").Font.Bold = True
        .Range("A4").value = "■ 約定一覧": .Range("A4").Font.Bold = True
        .Range("A2").value = "(デモ注文ログがありません)"
        .Columns.AutoFit
    End With

    '---- 6) Dashboard ボタン（UI_* ハンドラに統一）
    AddOrResetBtn wsD, "btnUpdateAll", "日次更新（PS→CSV）", "UI_UpdateAll_Click", "C2", 70, 20
  AddOrResetBtn wsD, "btnImport", "更新（CSV取込）", "UI_Import_Click", "I2", 70, 20
AddOrResetBtn wsD, "btnAutoStartDemo", "自動開始（デモ）", "UI_AutoStartDemo_Click", "D2", 70, 20
AddOrResetBtn wsD, "btnAutoStartLive", "自動開始（本番）", "UI_AutoStartLive_Click", "E2", 70, 20
AddOrResetBtn wsD, "btnAutoStop", "停止（パニック）", "UI_AutoStop_Click", "F2", 70, 20
AddOrResetBtn wsD, "btnOrders", "注文状況パネル更新", "UI_RefreshOrdersPanel_Click", "H2", 70, 20
'（日次PSボタンを使う場合）
AddOrResetBtn wsD, "btnPS", "日次PS 実行", "UI_RunPipeline_Click", "J2", 70, 20

AddOrResetBtn wsD, "btnSimTradeStart", "執行開始", "UI_SimTradeStart_Click", "K2", 70, 20
AddOrResetBtn wsD, "btnSimTradeStop", "執行停止", "UI_SimTradeStop_Click", "L2", 70, 20
AddOrResetBtn wsD, "btnSimTradeClear", "執行クリア", "UI_SimTradeClear_Click", "M2", 70, 20


    '---- 7) J/dJ/vEMA の数式を敷設（保護不要なら False、上書き不可にするなら True）
    InstallJKVFormulas False

    LogEvent "SYS", "SetupNewBook", "ready"
    SetStatus "ready"
    MsgBox "BookNew 完了：Dashboard/LOG/Orders を再生成し、J/dJ/vEMA の数式を敷設しました。", vbInformation
End Sub

' Dashboardにボタンを1つ作成（既存があれば削除→再作成）
Private Sub AddOrResetBtn(ws As Worksheet, ByVal nm As String, ByVal cap As String, _
                          ByVal macro As String, ByVal anchor As String, _
                          ByVal w As Single, ByVal h As Single)
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    With ws.Shapes.AddFormControl(xlButtonControl, ws.Range(anchor).Left, ws.Range(anchor).Top, w, h)
        .name = nm
        .TextFrame.Characters.Text = cap
        .OnAction = macro
    End With
End Sub


'----------------------------------------------------------
' 自動開始/停止（最小）
'----------------------------------------------------------
Public Sub AutoStart(ByVal demo As Boolean)
EnsureLogInfra:     EnsureDirs
    DEMO_MODE = demo: g_OrderID = 1
    ImportTopCSV
    ResolveRssLabels
    LiveStart
    LogEvent "SYS", "AutoStart", IIf(DEMO_MODE, "DEMO", "LIVE"): SetStatus IIf(demo, "自動開始（デモ）", "自動開始（本番）")
End Sub

Public Sub AutoStop()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC), "LiveTick", , False
    LogEvent "SYS", "LiveStop", "": LogEvent "SYS", "AutoStop", "": SetStatus "停止（パニック）"
End Sub

'----------------------------------------------------------
' Pipeline（非同期）※必要なときだけ使用
'----------------------------------------------------------
Public Sub BtnRunPipeline_Click()
    On Error GoTo FIN
    EnsureDirs
    Dim ps As String, script As String, log As String, wrap As String
    ps = GetPSPath(): If Len(ps) = 0 Then MsgBox "PowerShell が見つかりません。", vbExclamation: Exit Sub
    script = BASE_DIR & "\scripts\run_full.ps1"
    If Dir$(script) = "" Then MsgBox "run_full.ps1 が見つかりません: " & script, vbExclamation: Exit Sub
    log = LOG_DIR & "\pipeline_" & Format(Now, "yyyy-MM-dd_HHnnss") & ".txt"
    wrap = LOG_DIR & "\run_wrapper_" & Format(Now, "yyyy-MM-dd_HHnnss") & ".ps1"
    If Not WriteWrapper(script, log, wrap) Then MsgBox "PowerShell ラッパー作成に失敗", vbExclamation: Exit Sub
    RunPS_Async ps, wrap
    LogEvent "PS", "Pipeline launch", log
    SetStatus "起動: " & log
FIN:
End Sub

Private Function GetPSPath() As String
    If Dir$("C:\Program Files\PowerShell\7\pwsh.exe") <> "" Then GetPSPath = "C:\Program Files\PowerShell\7\pwsh.exe": Exit Function
    If Dir$("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe") <> "" Then GetPSPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe": Exit Function
    GetPSPath = ""
End Function

Private Function WriteWrapper(ByVal ps1 As String, ByVal log As String, ByVal wrap As String) As Boolean
    On Error GoTo FAIL
    Dim f As Integer: f = FreeFile
    Open wrap For Output As #f
    Print #f, "$ErrorActionPreference='Continue'"
    Print #f, "$log='" & Replace(log, "'", "''") & "'"
    Print #f, """START $(Get-Date -Format o)"" | Out-File -FilePath $log -Encoding utf8"
    Print #f, "try { & '" & Replace(ps1, "'", "''") & "' *>&1 | Tee-Object -FilePath $log -Append ; $code=$LASTEXITCODE }"
    Print #f, "catch { ""EXCEPTION: $($_)"" | Out-File -FilePath $log -Append ; $code = 1 }"
    Print #f, """END $(Get-Date -Format o) exit="" + $code | Out-File -FilePath $log -Append"
    Print #f, "exit $code"
    Close #f: WriteWrapper = True: Exit Function
FAIL:
    On Error Resume Next: Close #f: WriteWrapper = False
End Function

Private Sub RunPS_Async(ByVal psPath As String, ByVal wrapper As String)
    If Len(psPath) = 0 Or Len(wrapper) = 0 Then Exit Sub
    Dim cmd As String
    cmd = """" & psPath & """" & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File " & Chr$(34) & wrapper & Chr$(34)
    CreateObject("WScript.Shell").Run cmd, 7, False
End Sub

'========================  休場日シミュレーター  ========================
' Close/VWAPd をランダムウォークさせ、J/dJ/vEMA を連続更新します。
' 使い方：
'   Alt+F8 → SimStart_Fast（1秒/行50/ボラ0.10%） or  SimStart_Slow（3秒/行200/ボラ0.05%）
'   Alt+F8 → SimStop（停止） → SimReset（取り込み直後に戻す）
'======================================================================


'================= セクション丸ごと：SimStart_Fast/Slow/SimStart =================
Public Sub SimStart_Fast(): SimStart 0.001, 1, 50: End Sub
Public Sub SimStart_Slow(): SimStart 0.0005, 3, 200: End Sub

Public Sub SimStart(ByVal volPct As Double, ByVal stepSec As Long, ByVal rows As Long)
    CancelSim
    If stepSec <= 0 Then stepSec = 1
    If volPct <= 0 Then volPct = 0.001
    If rows <= 0 Then rows = 100

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    simVolPct = volPct: simStepSec = stepSec: simRows = rows
    Randomize Timer

    ' 重要：パス状態を初期化（前回の停止位置を引きずらない）
    If Not gSimPath Is Nothing Then gSimPath.RemoveAll

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then
        MsgBox "候補がありません。先に [更新（CSV取込）]。", vbExclamation
        Exit Sub
    End If

    Dim h As Range: Set h = rng.rows(1)
    Dim cC&, cVW&: cC = FindCol(h, "Close"): cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    Dim lastRow&, useN&, r&
    lastRow = ws.Cells(ws.rows.Count, cC).End(xlUp).row
    useN = Application.Min(rows, lastRow - (rng.row + 1) + 1)
    If useN <= 0 Then MsgBox "データ行がありません。", vbExclamation: Exit Sub

    ReDim simSnap(1 To useN, 1 To 2)
    For r = 1 To useN
        simSnap(r, 1) = d(ws.Cells(rng.row + r, cC).value)
        simSnap(r, 2) = d(IIf(cVW > 0, ws.Cells(rng.row + r, cVW).value, ws.Cells(rng.row + r, cC).value))
    Next r

    simRunning = True
    LogEvent "SIM", "Start", "vol=" & simVolPct & " step=" & simStepSec & "s rows=" & useN
    SimTick     ' 初回即時
End Sub
'================= セクション終わり：SimStart =================





'================= セクション丸ごと：SimStop =================
Public Sub SimStop()
    CancelSim
    simRunning = False
    LogEvent "SIM", "Stop", ""
End Sub
'================= セクション終わり：SimStop =================




Public Sub SimReset()
    If simRunning Then
        MsgBox "SimStop の後に実行してください。", vbInformation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then Exit Sub

    Dim h As Range: Set h = rng.rows(1)
    Dim cC&, cVW&: cC = FindCol(h, "Close"): cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")
    If cC = 0 Or UBound(simSnap, 1) = 0 Then Exit Sub

    Dim r As Long
    For r = 1 To UBound(simSnap, 1)
        If simSnap(r, 1) <> 0 Then ws.Cells(rng.row + r, cC).value = simSnap(r, 1)
        If cVW > 0 And simSnap(r, 2) <> 0 Then ws.Cells(rng.row + r, cVW).value = simSnap(r, 2)
    Next r
    LogEvent "SIM", "Reset", "rows=" & UBound(simSnap, 1)
End Sub
'================= セクション終わり：SimStart / Stop / Reset =================



'=================  セクション丸ごと差し替え：SimTick  =================
' ・SimStart_* からスケジュールされ、休場日でも Close/VWAPd を擬似的に更新
' ・J/dJ/vEMA は、J列が数式ならシート計算に、数値ならValue2で上書き
' ・価格更新後に SimTrade_OnTick を呼んで、ENTRY/EXIT～Orders記録まで行う
'======================================================================
'=================  セクション丸ごと：SimTick（OHLC滑らかパス）  =================
'================= セクション丸ごと：SimTick（診断＋OHLC滑らかパス） =================
'================= セクション丸ごと：SimTick（コード列なし耐性版） =================
Public Sub SimTick()
    On Error GoTo SCHEDULE
    If Not simRunning Then GoTo SCHEDULE

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then GoTo SCHEDULE

    ' ヘッダ
    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cOpen&, cHigh&, cLow&, cClose&, cVW&, cA&, cJ&, cdJ&, cvEMA&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "銘柄コード", "Code", "コード")
    cOpen = FindCol(h, "Open")
    cHigh = FindCol(h, "High")
    cLow = FindCol(h, "Low")
    cClose = FindCol(h, "Close")
    cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")
    cA = FindCol(h, "ATR5")
    cJ = FindCol(h, "J")
    cdJ = FindCol(h, "dJ")
    cvEMA = FindCol(h, "vEMA")

    Static tickCnt As Long: tickCnt = tickCnt + 1
    If cClose * cJ = 0 Then
        If tickCnt <= 3 Then LogEvent "SIMDBG", "HeaderMissing", "cClose=" & cClose & " cJ=" & cJ
        GoTo SCHEDULE
    End If
    If tickCnt = 1 Then LogEvent "SIMDBG", "Tick", "rngRows=" & rng.rows.Count

    ' 範囲
    Dim firstRow As Long, lastRow As Long, useN As Long
    firstRow = rng.row + 1
    lastRow = ws.Cells(ws.rows.Count, cClose).End(xlUp).row
    If lastRow < firstRow Then GoTo SCHEDULE
    useN = Application.Min(simRows, lastRow - firstRow + 1)

    ' 価格更新（OHLC滑らかパス）
    Dim r As Long, key$, px#, vw#, atr#, jp#, jn#, dj#, ve#
    For r = firstRow To firstRow + useN - 1
        ' コード列が無い場合は行番号で擬似キー
        If cT > 0 Then
            key = TickerToCode(CStr(ws.Cells(r, cT).value))
            If Len(key) <> 4 Then key = "R" & CStr(r)
        Else
            key = "R" & CStr(r)
        End If

        px = SimPath_NextPrice_OHLC(ws, r, cOpen, cHigh, cLow, cClose, cVW, key, vw)

        ' 書戻し
        ws.Cells(r, cClose).Value2 = px
        If cVW > 0 Then ws.Cells(r, cVW).Value2 = vw

        ' J/dJ/vEMA（数式セルは計算に任せる）
        If Not ws.Cells(r, cJ).HasFormula Then
            atr = d(IIf(cA > 0, ws.Cells(r, cA).value, 0)): If atr <= 0 Then atr = Abs(px) * 0.005
            jp = d(ws.Cells(r, cJ).value)
            jn = (px - IIf(cVW > 0, d(ws.Cells(r, cVW).value), px)) / atr
            dj = jn - jp
            ve = 0.3 * dj + 0.7 * d(IIf(cvEMA > 0, ws.Cells(r, cvEMA).value, 0))
            ws.Cells(r, cJ).Value2 = jn
            If cdJ > 0 Then ws.Cells(r, cdJ).Value2 = dj
            If cvEMA > 0 Then ws.Cells(r, cvEMA).Value2 = ve
        End If
    Next r

    If Application.Calculation = xlCalculationManual Then ws.Calculate

SCHEDULE:
    On Error Resume Next
    SimTrade_OnTick
    SetStatus "SIM @" & Format$(Now, "HH:NN:SS")
    ReschedSim
End Sub
'================= セクション終わり：SimTick =================





' 標準正規乱数（Box-Muller）
Private Function RandN() As Double
    Static spare As Double, hasSpare As Boolean
    Dim u1 As Double, u2 As Double
    If hasSpare Then
        RandN = spare: hasSpare = False
    Else
        Do
            u1 = Rnd: u2 = Rnd
        Loop While u1 <= 0#
        RandN = Sqr(-2 * log(u1)) * Cos(2 * WorksheetFunction.Pi() * u2)
        spare = Sqr(-2 * log(u1)) * Sin(2 * WorksheetFunction.Pi() * u2)
        hasSpare = True
    End If
End Function


'=== 選択中のデータ行の J/dJ/vEMA を RSS 値で強制更新 ===
Public Sub ForceUpdateJ_ActiveRow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then MsgBox "候補がありません": Exit Sub

    Dim r As Long: r = ActiveCell.row
    If r <= rng.row Then r = rng.row + 1      ' 見出し行のときは 1 行下

    ForceUpdateJ_Core r
End Sub

'=== 先頭データ行（B5の 1 行下）を強制更新 ===
Public Sub ForceUpdateJ_FirstDataRow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then MsgBox "候補がありません": Exit Sub
    ForceUpdateJ_Core rng.row + 1
End Sub

'=== 実体：行番号 r の J/dJ/vEMA を RSS→式で更新（LOGに詳細を出力） ===
Private Sub ForceUpdateJ_Core(ByVal r As Long)
    On Error GoTo FAIL

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then
        LogEvent "FORCEJ", "NoData", "rng is empty": Exit Sub
    End If

    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cvEMA&, cVW&, tit As String
    cT = FindColAny(h, tit, "Ticker", "code", "銘柄コード", "Code")
    cC = FindCol(h, "Close"): cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J")
    cdJ = FindCol(h, "dJ"):   cvEMA = FindCol(h, "vEMA")
    cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    If cC * cA * cJ = 0 Then
        LogEvent "FORCEJ", "HeaderMissing", "cT=" & cT & "(" & tit & "), cC=" & cC & ", cA=" & cA & ", cJ=" & cJ
        Exit Sub
    End If

    ' 範囲チェック
    If r < rng.row + 1 Or r > rng.row + rng.rows.Count - 1 Then
        LogEvent "FORCEJ", "RowOut", "r=" & r: Exit Sub
    End If

    ' コード取得（常に文字列化→銘柄コード化）
    Dim tkr As String, code As String
    tkr = Trim$(CStr(rng.Cells(r, cT).value))
    code = TickerToCode(tkr)
    If Len(code) <> 4 Then
        LogEvent "FORCEJ", "BadCode", "tkr=" & tkr & " code=" & code
        Exit Sub
    End If

    ' RSS 取得（Variant→安全数値化）
    Dim nPx As String, nVw As String
    Dim vpx As Variant, vVW As Variant
    vpx = RSS_Get(code, "PRICE_CAND", nPx)
    vVW = RSS_Get(code, "VWAP_CAND", nVw)

    Dim px As Double, vw As Double, atr As Double
    If IsNumeric(vpx) Then px = CDbl(vpx) Else px = 0#
    If IsNumeric(vVW) Then vw = CDbl(vVW) Else vw = 0#
    atr = d(rng.Cells(r, cA).value)

    LogEvent "FORCEJ", "px/vw/atr", code & _
        " px=" & Format$(px, "0.###") & "(" & nPx & ")" & _
        " vw=" & Format$(vw, "0.###") & "(" & nVw & ")" & _
        " atr=" & Format$(atr, "0.###")

    If px <= 0# Or vw <= 0# Or atr <= 0# Then
        MsgBox "px/vw/atr のいずれかが 0 で更新できません。", vbExclamation
        Exit Sub
    End If

    ' J/dJ/vEMA 計算→書込み
    Dim jp As Double, jn As Double, dj As Double, ve As Double
    jp = d(rng.Cells(r, cJ).value)
    jn = (px - vw) / atr
    dj = jn - jp
    ve = 0.3 * dj + 0.7 * d(IIf(cvEMA > 0, rng.Cells(r, cvEMA).value, 0))

 ' --- FORCEJ 書込み部（ここから置換） ---
rng.Cells(r, cC).Value2 = px
If cVW > 0 Then rng.Cells(r, cVW).Value2 = vw

If rng.Cells(r, cJ).HasFormula Then rng.Cells(r, cJ).Value2 = rng.Cells(r, cJ).Value2
rng.Cells(r, cJ).Value2 = jn
If cdJ > 0 Then rng.Cells(r, cdJ).Value2 = dj
If cvEMA > 0 Then rng.Cells(r, cvEMA).Value2 = ve
' --- FORCEJ 書込み部（置換ここまで） ---


    LogEvent "FORCEJ", "Jupdated", code & " " & Format$(jp, "0.###") & "→" & Format$(jn, "0.###")
    Exit Sub
FAIL:
    LogEvent "ERR", "FORCEJ", Err.Description
    MsgBox "FORCEJ でエラー: " & Err.Description, vbExclamation
End Sub




'=== Dashboard の「日次PS 実行」ボタン重複の解消＆再配置 ===
Public Sub FixPipelineButton()
    Dim ws As Worksheet, shp As Shape, hit As Long
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)

    ' 1) 既存の「日次PS 実行」っぽいボタンを全部消す（キャプション/OnAction両面で判定）
    For Each shp In ws.Shapes
        On Error Resume Next
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                Dim cap$: cap = ""
                cap = shp.TextFrame.Characters.Text
                If cap Like "*日次PS*" Or shp.OnAction = "UI_RunPipeline_Click" Or shp.OnAction = "BtnRunPipeline_Click" Then
                    shp.Delete
                End If
            End If
        End If
        On Error GoTo 0
    Next shp

    ' 2) 正規のPSボタンを1つだけ配置（位置は J2：必要なら変えてOK）

AddOrResetBtn ws, "btnPS", "日次PS 実行", "UI_RunPipeline_Click", "J2", 120, 28


    LogEvent "FIX", "PipelineButton", "Reinstalled"
End Sub

'=== J / dJ / d2J / vEMA の数式を敷設（全行・統一） ===
'================= セクション丸ごと差し替え：InstallJKVFormulas（シード版） =================
Public Sub InstallJKVFormulas(Optional ByVal protectCols As Boolean = False)
    On Error GoTo FIN

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then Exit Sub

    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cD2&, cvEMA&, cVW&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "銘柄コード", "Code")
    cC = FindCol(h, "Close")
    cA = FindCol(h, "ATR5")
    cJ = FindCol(h, "J")
    cdJ = FindCol(h, "dJ")
    cD2 = FindCol(h, "d2J")
    cvEMA = FindCol(h, "vEMA")
    cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    If cC * cA * cJ = 0 Then Exit Sub

    Dim firstRow As Long, lastRow As Long
    firstRow = rng.row + 1
    lastRow = ws.Cells(ws.rows.Count, cC).End(xlUp).row
    If lastRow < firstRow Then Exit Sub

    ' R1C1 オフセット
    Dim offC&, offVW&, offATR&: offC = cC - cJ: offVW = cVW - cJ: offATR = cA - cJ

    ' J: (Close - VWAPd) / ATR5
    Dim fJ$: fJ = "=(RC[" & offC & "]-RC[" & offVW & "])/RC[" & offATR & "]"
    ws.Range(ws.Cells(firstRow, cJ), ws.Cells(lastRow, cJ)).FormulaR1C1 = fJ

    ' dJ: 先頭= Jr - J(r+1) / 以降=差分
    If cdJ > 0 Then
        Dim offJ_from_dJ&: offJ_from_dJ = cJ - cdJ
        ws.Cells(firstRow, cdJ).FormulaR1C1 = "=RC[" & offJ_from_dJ & "]-R[1]C[" & offJ_from_dJ & "]"
        If firstRow + 1 <= lastRow Then
            ws.Range(ws.Cells(firstRow + 1, cdJ), ws.Cells(lastRow, cdJ)).FormulaR1C1 = _
                "=RC[-" & (cdJ - cJ) & "]-R[-1]C"
        End If
    End If

    ' d2J: 先頭= dJr - dJ(r+1) / 以降=差分
    If cD2 > 0 And cdJ > 0 Then
        Dim off_dJ_from_d2&: off_dJ_from_d2 = cdJ - cD2
        ws.Cells(firstRow, cD2).FormulaR1C1 = "=RC[" & off_dJ_from_d2 & "]-R[1]C[" & off_dJ_from_d2 & "]"
        If firstRow + 1 <= lastRow Then
            ws.Range(ws.Cells(firstRow + 1, cD2), ws.Cells(lastRow, cD2)).FormulaR1C1 = _
                "=RC[-" & (cD2 - cdJ) & "]-R[-1]C"
        End If
    End If

    ' vEMA: 先頭= dJ（K列）/ 以降= 0.3*dJ + 0.7*前回
    If cvEMA > 0 And cdJ > 0 Then
        Dim off_dJ_from_vEMA&: off_dJ_from_vEMA = cdJ - cvEMA
        ws.Cells(firstRow, cvEMA).FormulaR1C1 = "=RC[" & off_dJ_from_vEMA & "]"
        If firstRow + 1 <= lastRow Then
            ws.Range(ws.Cells(firstRow + 1, cvEMA), ws.Cells(lastRow, cvEMA)).FormulaR1C1 = _
                "=0.3*RC[" & (cdJ - cvEMA) & "]+0.7*R[-1]C"
        End If
    End If

    If protectCols Then
        With ws
            .Unprotect
            If cJ > 0 Then .Columns(cJ).Locked = True
            If cdJ > 0 Then .Columns(cdJ).Locked = True
            If cD2 > 0 Then .Columns(cD2).Locked = True
            If cvEMA > 0 Then .Columns(cvEMA).Locked = True
            .Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
        End With
    End If

    LogEvent "FORMULA", "InstallJKV", "rows=" & (lastRow - firstRow + 1) & _
        " J=" & cJ & " dJ=" & cdJ & " d2J=" & cD2 & " vEMA=" & cvEMA
FIN:
End Sub
'=================  セクション丸ごと差し替え：InstallJKVFormulas（シード版）ここまで  =================



'================= セクション丸ごと：UI（ボタン） =================
Public Sub UI_SimTradeStart_Click():  SimTradeStart_Engine:  End Sub
Public Sub UI_SimTradeStop_Click():   SimTradeStop_Engine:   End Sub
Public Sub UI_SimTradeClear_Click():  SimTradeClear_Engine:  End Sub
'================= セクション終わり：UI =================

'============= セクション丸ごと：疑似執行 Start/Stop/Clear(Engine) =============
Public Sub SimTradeStart_Engine()
    If gSimPos Is Nothing Then Set gSimPos = CreateObject("Scripting.Dictionary")
    simTradeOn = True
    LogEvent "SIMTRADE", "Start", "TH=" & SIM_TH_ABSJ & _
             " TP=" & SIM_K_TP_ATR & " SL=" & SIM_K_SL_ATR & " TMAX=" & SIM_TMAX_BAR
End Sub

Public Sub SimTradeStop_Engine()
    simTradeOn = False
    LogEvent "SIMTRADE", "Stop", ""
End Sub

Public Sub SimTradeClear_Engine()
    If Not gSimPos Is Nothing Then gSimPos.RemoveAll
    LogEvent "SIMTRADE", "Clear", "positions=0"
End Sub
'============= セクション終わり：Start/Stop/Clear(Engine) =============

'================= ここからはコメントのみ（解説） ==================
' ・この下には実行コードを置かないこと（End Subの後はコメントのみ許可）
' ・SimTickの末尾で SimTrade_OnTick を呼ぶ想定
' ・Orders への記録は SimLogEntry / SimLogExit を使用
'===================================================================


'=================  セクション丸ごと差し替え：SimTrade_OnTick ここまで  =================

'=================  セクション丸ごと差し替え：SimTrade_OnTick ここまで  =================


'==== Orders への記録（ENTRY）6引数版 ====
Private Sub SimLogEntry(ByVal code As String, ByVal side As String, _
                        ByVal qty As Long, ByVal px As Double, _
                        ByVal tp As Double, ByVal sl As Double)

    Dim ws As Worksheet
    If SheetExists(ORD_SHEET) = False Then
        Set ws = ThisWorkbook.Worksheets.Add: ws.name = ORD_SHEET
    Else
        Set ws = ThisWorkbook.Worksheets(ORD_SHEET)
    End If

    ' ヘッダ初期化
    If Trim$(ws.Range("A1").value) <> "■ 注文一覧" Then
        ws.Cells.Clear
        ws.Range("A1").value = "■ 注文一覧": ws.Range("A1").Font.Bold = True
        ws.Range("A4").value = "■ 約定一覧": ws.Range("A4").Font.Bold = True
        ws.Range("A2").value = "(デモ注文ログがありません)"
    End If

    ' 行追加
    Dim row As Long: row = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If row < 2 Then row = 1
    If ws.Range("A2").value = "(デモ注文ログがありません)" Then ws.Range("A2").ClearContents: row = 1

    With ws
        .Cells(row + 1, 1).value = Format$(Now, "yyyy-MM-dd HH:mm:ss")
        .Cells(row + 1, 2).value = "SIM-ENTRY"
        .Cells(row + 1, 3).value = side & " " & code & " x" & qty & " @" & Format$(px, "0.000")
        .Cells(row + 1, 4).value = "tp=" & Format$(tp, "0.000") & " sl=" & Format$(sl, "0.000")
    End With
End Sub

Private Sub SimLogExit(ByVal code As String, ByVal side As String, _
                       ByVal qty As Long, ByVal px As Double, _
                       ByVal reason As String, ByVal pnl As Double, _
                       ByVal bars As Long)
    Dim ws As Worksheet
    If SheetExists(ORD_SHEET) = False Then
        Set ws = ThisWorkbook.Worksheets.Add: ws.name = ORD_SHEET
    Else
        Set ws = ThisWorkbook.Worksheets(ORD_SHEET)
    End If

    Dim row As Long: row = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If row < 2 Then row = 1

    With ws
        .Cells(row + 1, 1).value = Format$(Now, "yyyy-MM-dd HH:mm:ss")
        .Cells(row + 1, 2).value = "SIM-EXIT"
        .Cells(row + 1, 3).value = side & " " & code & " x" & qty & " @" & Format$(px, "0.000")
        .Cells(row + 1, 4).value = "reason=" & reason & " pnl=" & Format$(pnl, "0.000") & " bars=" & bars
    End With
End Sub


'=================  セクション追加：SimTrade_OnTick + 起動/停止  =================
' 設定値（未宣言なら先頭に宣言）
'Private Const SIM_TH_ABSJ   As Double = 0.8
'Private Const SIM_K_TP_ATR  As Double = 2#
'Private Const SIM_K_SL_ATR  As Double = 2#
'Private Const SIM_TMAX_BAR  As Long   = 60
'Private Const SIM_MIN_REENTRY As Long = 20
'Private Const SIM_QTY       As Long   = 100
'Private simTradeOn As Boolean
'Private gSimPos As Object   ' Scripting.Dictionary



'=================  セクション丸ごと：SimTrade_OnTick（完全版）  =================
' ・SimTick が Close/VWAPd を更新した“後”に呼ぶ（SimTick末尾 SCHEDULE: で呼ぶ）
' ・EXIT → ENTRY の順で処理し、Orders に SIM-EXIT / SIM-ENTRY を追記
' ・LOG に Digest を出し、発注が出ない内訳を可視化（price0/atrAdj/hold/cool/noSignal/badCode）
Public Sub SimTrade_OnTick()

    If Not simTradeOn Then Exit Sub
    On Error GoTo FIN

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then Exit Sub

    '--- ヘッダ検出（表記ゆらぎ対応） ---
    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cD2&, cvEMA&, cVW&, tit As String
    cT = FindColAny(h, tit, "Ticker", "code", "銘柄コード", "Code")
    cC = FindCol(h, "Close")
    cA = FindCol(h, "ATR5")
    cJ = FindCol(h, "J")
    cdJ = FindCol(h, "dJ")
    cD2 = FindCol(h, "d2J")
    cvEMA = FindCol(h, "vEMA")
    cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    If cT * cC * cA * cJ = 0 Then
        LogEvent "SIMTRADE", "HeaderMissing", _
                 "RowCnt=" & rng.rows.Count & " cT=" & cT & "(" & tit & ")" & _
                 " cC=" & cC & " cA=" & cA & " cJ=" & cJ
        Exit Sub
    End If

    '--- 走査範囲 ---
    Dim firstRow As Long, lastRow As Long
    firstRow = rng.row + 1
    lastRow = ws.Cells(ws.rows.Count, cC).End(xlUp).row
    If lastRow < firstRow Then Exit Sub

    If gSimPos Is Nothing Then Set gSimPos = CreateObject("Scripting.Dictionary")

    '============================== 1) EXIT 先 ==============================
    Dim k As Variant, pos As Object
    Dim px#, atr#, exitNow As Boolean, reason$
    If gSimPos.Count > 0 Then
        Dim ks As Variant: ks = gSimPos.Keys
        Dim i As Long
        For i = LBound(ks) To UBound(ks)
            If gSimPos.Exists(ks(i)) Then
                Set pos = gSimPos(ks(i))
                Dim r As Long: r = pos("row")
                px = d(ws.Cells(r, cC).value)
                atr = d(ws.Cells(r, cA).value): If atr <= 0 Then atr = 0.0001

                exitNow = False: reason = ""
                If pos("side") = "BUY" Then
                    If px >= pos("tp") Then exitNow = True: reason = "TP"
                    If px <= pos("sl") Then
                        exitNow = True
                        If Len(reason) > 0 Then reason = reason & "+SL" Else reason = "SL"
                    End If
                Else
                    If px <= pos("tp") Then exitNow = True: reason = "TP"
                    If px >= pos("sl") Then
                        exitNow = True
                        If Len(reason) > 0 Then reason = reason & "+SL" Else reason = "SL"
                    End If
                End If

                ' 時間切れ
                pos("bars") = pos("bars") + 1
                If pos("bars") >= SIM_TMAX_BAR Then
                    exitNow = True
                    If Len(reason) > 0 Then reason = reason & "+TMAX" Else reason = "TMAX"
                End If

                If exitNow Then
                    Dim pnl#
                    If pos("side") = "BUY" Then
                        pnl = (px - pos("px")) * pos("qty")
                    Else
                        pnl = (pos("px") - px) * pos("qty")
                    End If
                    SimLogExit pos("code"), pos("side"), pos("qty"), px, reason, pnl, pos("bars")
                    pos("open") = False: pos("cool") = SIM_MIN_REENTRY
                    gSimPos(pos("code")) = pos
                Else
                    If pos.Exists("cool") Then If pos("cool") > 0 Then pos("cool") = pos("cool") - 1
                    gSimPos(pos("code")) = pos
                End If
            End If
        Next i

        ' 完全クローズ（open=False かつ cool<=0）を除去
        Dim del As New Collection
        For Each k In gSimPos.Keys
            Set pos = gSimPos(k)
            If (Not pos("open")) And pos("cool") <= 0 Then del.Add k
        Next k
        For Each k In del: gSimPos.Remove k: Next k
    End If

    '============================== 2) ENTRY ==============================
    Dim row As Long, code$, j#, dj#, ve#, side$, tp#, sl#, qty&, vw#
    Dim px0&, atr0&, sig0&, hold0&, cd0&  ' digest 用カウンタ
    qty = SIM_QTY

    For row = firstRow To lastRow
        code = TickerToCode(CStr(ws.Cells(row, cT).value))
        If Len(code) <> 4 Then cd0 = cd0 + 1: GoTo NextRow

        j = d(ws.Cells(row, cJ).value)
        dj = d(ws.Cells(row, cdJ).value)
        ve = d(ws.Cells(row, cvEMA).value)
        px = d(ws.Cells(row, cC).value)
        vw = d(IIf(cVW > 0, ws.Cells(row, cVW).value, px))
        If px <= 0 Then px0 = px0 + 1: GoTo NextRow

        atr = d(ws.Cells(row, cA).value)
        If atr <= 0 Then atr = 0.0001: atr0 = atr0 + 1

        ' 保有/クールダウン中はスキップ
        If gSimPos.Exists(code) Then
            Set pos = gSimPos(code)
            If pos("open") Or pos("cool") > 0 Then hold0 = hold0 + 1: GoTo NextRow
        End If

        ' 逆張りシグナル
        If Abs(j) >= SIM_TH_ABSJ And Sgn(ve) <> Sgn(dj) Then
            side = IIf(j < 0, "BUY", "SELL")
            If side = "BUY" Then
                tp = px + SIM_K_TP_ATR * atr: sl = px - SIM_K_SL_ATR * atr
            Else
                tp = px - SIM_K_TP_ATR * atr: sl = px + SIM_K_SL_ATR * atr
            End If

            ' ポジション保存
            Set pos = CreateObject("Scripting.Dictionary")
            pos("open") = True: pos("code") = code: pos("side") = side
            pos("qty") = qty: pos("px") = px: pos("tp") = tp: pos("sl") = sl
            pos("row") = row: pos("bars") = 0: pos("cool") = 0
            gSimPos(code) = pos

            ' Ordersへ記録
            SimLogEntry code, side, qty, px, tp, sl
        Else
            sig0 = sig0 + 1
        End If

NextRow:
    Next row

    ' ダイジェスト（発注が出ないときの内訳）
    LogEvent "SIMTRADE", "Digest", _
        "Rows=" & (lastRow - firstRow + 1) & _
        " price0=" & px0 & " atrAdj=" & atr0 & " hold/cool=" & hold0 & " noSignal=" & sig0 & " badCode=" & cd0

FIN:
    Exit Sub
End Sub
'=================  セクション丸ごと：SimTrade_OnTick（完全版）ここまで  =================


'=================  セクション追加：SimTrade_OnTick ここまで  =================

'================= セクション追加：UpdateAll（PS→Import→式→ラベル） =================
Public Sub UI_UpdateAll_Click()
    UpdateAll_Run 300   ' 最大待機300秒（環境に合わせて調整）
End Sub

'================= セクション丸ごと：UpdateAll_Run（PS失敗でも続行） =================
Private Sub UpdateAll_Run(ByVal maxWaitSec As Long)
    On Error GoTo FAIL
    EnsureDirs

    Dim ps$, script$, log$, wrap$, exitCode&
    ps = GetPSPath(): script = BASE_DIR & "\scripts\run_full.ps1"
    log = LOG_DIR & "\pipeline_" & Format$(Now, "yyyy-MM-dd_HHnnss") & ".txt"
    wrap = LOG_DIR & "\run_wrapper_" & Format$(Now, "yyyy-MM-dd_HHnnss") & ".ps1"

    If Dir$(ps) <> "" And Dir$(script) <> "" And WriteWrapper(script, log, wrap) Then
        exitCode = RunPS_Sync(ps, wrap, maxWaitSec)
        LogEvent "PS", "Pipeline sync", "exit=" & exitCode & " | " & log
        ' 失敗しても処理継続（MsgBoxは出さない）
    Else
        LogEvent "PS", "Pipeline skip", "ps/script/Wrapper NG"
    End If

    ImportTopCSV
    InstallJKVFormulas False
    ResolveRssLabels
    SetStatus "UpdateAll 完了"
    Exit Sub
FAIL:
    LogEvent "ERR", "UpdateAll", Err.Description
    ' 続行しない
End Sub
'================= セクション終わり：UpdateAll_Run =================


' 同期実行：PSを待機し、最大 maxWaitSec 秒でタイムアウト。戻り値は exit code（-1=タイムアウト）
Private Function RunPS_Sync(ByVal psPath As String, ByVal wrapper As String, ByVal maxWaitSec As Long) As Long
    On Error GoTo FAIL
    Dim cmd$, sh As Object, t As Single, p As Object
    cmd = """" & psPath & """" & " -NoProfile -ExecutionPolicy Bypass -File " & Chr$(34) & wrapper & Chr$(34)
    Set sh = CreateObject("WScript.Shell")
    Set p = sh.Exec(cmd)
    t = Timer
    Do While p.status = 0
        DoEvents
        If maxWaitSec > 0 And (Timer - t) > maxWaitSec Then
            RunPS_Sync = -1
            Exit Function
        End If
    Loop
    RunPS_Sync = p.exitCode
    Exit Function
FAIL:
    RunPS_Sync = -2
End Function
'================= セクション追加：UpdateAll ここまで =================




'================= セクション追加：手動キック（診断用） =================
Public Sub UI_SimKick_Click()
    LogEvent "SIMDBG", "Kick", "manual"
    SimTick
End Sub
'================= セクション追加ここまで =================




'================= セクション：Diag_WriteTest =================
Public Sub Diag_WriteTest()
    On Error GoTo FAIL
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    Dim h As Range:      Set h = rng.rows(1)
    Dim cC As Long:      cC = FindCol(h, "Close")
    If cC = 0 Then LogEvent "DIAG", "NoClose", "": Exit Sub
    Dim r As Long:       r = rng.row + 1
    Dim before As Double: before = d(ws.Cells(r, cC).value)
    ws.Cells(r, cC).Value2 = before * (1 + 0.001)        ' +0.1%
    LogEvent "DIAG", "WriteOK", "row=" & r & " colClose=" & cC & " " & before & " -> " & ws.Cells(r, cC).value
    Exit Sub
FAIL:
    LogEvent "DIAG", "WriteFAIL", Err.Description
End Sub
'================= セクション終わり：Diag_WriteTest =================

'================= セクション：SimTick_Minimal =================
Public Sub SimTick_Minimal()
    On Error GoTo SCHEDULE
    If Not simRunning Then GoTo SCHEDULE

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then GoTo SCHEDULE

    Dim h As Range: Set h = rng.rows(1)
    Dim cC As Long, cVW As Long, cA As Long, cJ As Long, cdJ As Long, cvE As Long
    cC = FindCol(h, "Close")
    cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")
    cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J"): cdJ = FindCol(h, "dJ"): cvE = FindCol(h, "vEMA")
    If cC = 0 Then GoTo SCHEDULE

    Dim firstRow As Long, lastRow As Long, r As Long, useN As Long
    firstRow = rng.row + 1
    lastRow = ws.Cells(ws.rows.Count, cC).End(xlUp).row
    If lastRow < firstRow Then GoTo SCHEDULE
    useN = Application.Min(IIf(simRows > 0, simRows, 50), lastRow - firstRow + 1)

    For r = firstRow To firstRow + useN - 1
        Dim px As Double, vw As Double, atr As Double, jp As Double, jn As Double, dj As Double, ve As Double
        px = d(ws.Cells(r, cC).value)
        If px <= 0# Then px = 100#
        ' ランダムウォーク
        px = px * (1 + RandN() * 0.001)
        ws.Cells(r, cC).Value2 = px

        If cVW > 0 Then
            vw = d(ws.Cells(r, cVW).value): If vw <= 0# Then vw = px
            vw = vw + 0.15 * (px - vw)           ' 簡易VWAP追随
            ws.Cells(r, cVW).Value2 = vw
        Else
            vw = px
        End If

        If cJ > 0 Then
            atr = d(IIf(cA > 0, ws.Cells(r, cA).value, 0#)): If atr <= 0# Then atr = Abs(px) * 0.005
            jp = d(ws.Cells(r, cJ).value)
            jn = (px - vw) / atr
            dj = jn - jp
            ve = 0.3 * dj + 0.7 * d(IIf(cvE > 0, ws.Cells(r, cvE).value, 0#))
            If Not ws.Cells(r, cJ).HasFormula Then ws.Cells(r, cJ).Value2 = jn
            If cdJ > 0 And Not ws.Cells(r, cdJ).HasFormula Then ws.Cells(r, cdJ).Value2 = dj
            If cvE > 0 And Not ws.Cells(r, cvE).HasFormula Then ws.Cells(r, cvE).Value2 = ve
        End If
    Next r

SCHEDULE:
    On Error Resume Next
    ReschedSim
End Sub
'================= セクション終わり：SimTick_Minimal =================

'================= セクション：Diag_UnprotectAll =================
Public Sub Diag_UnprotectAll()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ThisWorkbook.Unprotect
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
        ws.EnableCalculation = True
    Next ws
    On Error GoTo 0
End Sub
'================= 終わり =================

'================= セクション：Diag_ReportHeaders =================
Public Sub Diag_ReportHeaders()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Then LogEvent "DIAG", "NoRange", START_CELL: Exit Sub
    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cO&, cH&, cL&, cC&, cVW&, cA&, cJ&, cdJ&, cvE&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "銘柄コード", "Code", "コード")
    cO = FindCol(h, "Open"): cH = FindCol(h, "High"): cL = FindCol(h, "Low")
    cC = FindCol(h, "Close"): cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")
    cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J"): cdJ = FindCol(h, "dJ"): cvE = FindCol(h, "vEMA")
    LogEvent "DIAG", "Headers", _
      "start=" & START_CELL & " firstRow=" & rng.row + 1 & _
      " cT=" & cT & "(" & tit & ")" & " cC=" & cC & " cVW=" & cVW & _
      " cA=" & cA & " cJ=" & cJ & " cdJ=" & cdJ & " cvEMA=" & cvE
End Sub
'================= 終わり =================


'================= セクション：Diag_HardWriteLoop =================
Public Sub Diag_HardWriteLoop()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim tgt As Range:    Set tgt = ws.Range("Z2")   ' どこでも良い空きセル
    Dim i As Long
    For i = 1 To 10
        tgt.Value2 = "PING " & Format$(Now, "HH:NN:SS")
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
    LogEvent "DIAG", "HardWriteLoop", "done @" & Format$(Now, "HH:NN:SS")
End Sub
'================= 終わり =================

