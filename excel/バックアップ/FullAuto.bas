Attribute VB_Name = "FullAuto"
Option Explicit

'================= �Z�N�V�����F�錾�u���b�N�i���W���[�����x����1�񂾂��j =================
' ---- OHLC�p�X ----
Public Const SIMPATH_STEPS      As Long = 600
Public Const SIMPATH_VWAP_ALPHA As Double = 0.15
Public Const SIMPATH_NOISE      As Double = 0.002
Public gSimPath As Object                                  ' code �� state(Dictionary)

' ---- SIM ���s�G���W���̐ݒ�l ----
Public Const SIM_TH_ABSJ     As Double = 0.8
Public Const SIM_K_TP_ATR    As Double = 2#
Public Const SIM_K_SL_ATR    As Double = 2#
Public Const SIM_TMAX_BAR    As Long = 60
Public Const SIM_MIN_REENTRY As Long = 20
Public Const SIM_QTY         As Long = 100

' ---- ��ԁi�N���t���O�^�|�W�V�����^���O�j ----
Public simTradeOn As Boolean
Public gSimPos    As Object                              ' Scripting.Dictionary
Public LiveBusy   As Boolean
Public CsvBusy    As Boolean
Private CompletedLogs As Object
Private Const LOG_MAX As Long = 1000

' ---- RSS �������x���^�w�K�ς݃��x�� ----
Private Const PRICE_LABEL As String = "���ݒl"
Private Const VWAP_LABEL  As String = "�o�������d����"
Private gPriceLabel As String
Private gVwapLabel  As String

' ---- �x����V�~�����[�^��� ----
Private simRunning As Boolean
Private simNext    As Date
Private simStepSec As Long
Private simRows    As Long
Private simVolPct  As Double
Private simSnap()  As Double

' ---- �p�X�^�V�[�g ----
Public Const BASE_DIR    As String = "C:\AI\asagake"
Public Const SIGNAL_DIR  As String = BASE_DIR & "\data\signals\"
Public Const LIVE_DIR    As String = BASE_DIR & "\data\live"
Public Const LOG_DIR     As String = BASE_DIR & "\logs"
Public Const DASH_SHEET  As String = "Dashboard"
Public Const ORD_SHEET   As String = "Orders"
Public Const LOG_SHEET   As String = "LOG"
Public Const START_CELL  As String = "B5"

' ---- Live ����/������ ----
Public Const LIVE_INTERVAL_SEC As Long = 3
Public Const LIVE_J_MIN        As Double = 0.5

' ---- �f��/����/�M�p ----
Public DEMO_MODE   As Boolean
Public MARGIN_MODE As Boolean
Public Const LOT_DEFAULT As Long = 100

' ---- RSS���� ----
Public Const SOR_KBN  As String = "0"
Public Const ACC_KOZA As String = "0"
Public g_OrderID      As Long
'================= �Z�N�V�����F�錾�u���b�N�i�����܂Łj =================
'================= �Z�N�V�����FOnTime�X�P�W���[���i���i�t�H�[���o�b�N�j =================
' ���̃u���b�N�͐錾�u���b�N�̒����1�񂾂��z�u
'================= �Z�N�V�����FOnTime�X�P�W���[���i�ŏ��łփ��[�e�B���O�j =================
Private Function ProcPatterns() As Variant
    Dim b$: b = ThisWorkbook.name
    ' �� SimTick_Minimal ��D��Ăяo��
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
'================= �Z�N�V�����I��� =================



'==========================================================
' SHINSOKU ? ����Ńt�����W���[��
' �ESignals�̎��������i������������΍ŐV���j
' �ECSV��荞�݁iUTF-8/�_�u���N�H�[�g�Ή��̌��S�p�[�T�j
' �ETicker�񖼂̗h��z���iTicker/code/�����R�[�h/Code�j
' �ERssMarket�F�������x���u���ݒl�v�u�o�������d���ρv���ŗD�恨�ی����x����������
' �ELive(3�b)�� J/dJ/vEMA ��K���X�V�i�l����ꂽ�s�̂݁j
' �E�����N���b�N���� LiveOnce_Debug
' �E�f�������FCSV�ۑ��{Orders�ɑ��ǋL
' �EBookNew�F���S���Z�b�g�{�{�^���Ċ���
' �EUI�n���h���i�d�����������ӂɁj
' �E����PS�iPipeline�j�N���p�FBtnRunPipeline_Click + PS���[�e�B���e�B
'==========================================================

'================= �Z�N�V�����FSmoothStep�i�u���j =================
Private Function SmoothStep(ByVal u As Double) As Double
    If u < 0# Then
        u = 0#
    ElseIf u > 1# Then
        u = 1#
    End If
    SmoothStep = u * u * (3# - 2# * u)
End Function
'================= �Z�N�V�����FSmoothStep�i�����܂Łj =================




' OHLC �ɉ����āuOpen��(High/Low)��(Low/High)��Close�v��ʉ߂��銊�炩�p�X
' �Ecode���Ƃɏ�Ԃ�ێ��igSimPath: Dictionary�j
' �E�߂�l�F���� Close �l�Avwap �� ByRef �ŕԂ�
'================= �Z�N�V�����FSimPath_NextPrice_OHLC�i�[��HL�с{Seg�i�s�j =================
Private Function SimPath_NextPrice_OHLC(ByVal ws As Worksheet, ByVal r As Long, _
    ByVal cOpen As Long, ByVal cHigh As Long, ByVal cLow As Long, ByVal cClose As Long, _
    ByVal cVW As Long, ByVal code As String, ByRef vwap As Double) As Double

    If gSimPath Is Nothing Then Set gSimPath = CreateObject("Scripting.Dictionary")

    Dim o#, h#, l#, c#
    o = d(IIf(cOpen > 0, ws.Cells(r, cOpen).value, ws.Cells(r, cClose).value))
    h = d(IIf(cHigh > 0, ws.Cells(r, cHigh).value, o))
    l = d(IIf(cLow > 0, ws.Cells(r, cLow).value, o))
    c = d(ws.Cells(r, cClose).value): If c <= 0 Then c = o

    ' HL�������Ȃ�[�������W�𐶐��i�}max(0.05%, simVolPct)�j
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

    ' 1�e�B�b�N�i�߂�
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

    ' ���Z�O�����g�֐i�s
    If t >= tA Then
        If seg < 3 Then seg = seg + 1
    End If

    ' �����m�C�Y�{�N�����v
    price = price * (1 + RandN() * SIMPATH_NOISE)
    If price < l Then price = l
    If price > h Then price = h

    ' VWAP �ǐ�
    vwap = vwap + SIMPATH_VWAP_ALPHA * (price - vwap)

    ' �ۑ�
    state("t") = t: state("seg") = seg: state("vwap") = vwap
    gSimPath(code) = state

    SimPath_NextPrice_OHLC = price
End Function
'================= �Z�N�V�����I���FSimPath_NextPrice_OHLC =================




'=== UI handlers�i�{�^���p�B���O�Փ˂��Ȃ��悤 UI_ �ړ����ɓ���j ===
Public Sub UI_Import_Click():             ImportTopCSV: End Sub
Public Sub UI_AutoStartDemo_Click():      AutoStart True: End Sub
Public Sub UI_AutoStartLive_Click():      AutoStart False: End Sub
Public Sub UI_AutoStop_Click():           AutoStop: End Sub
Public Sub UI_RefreshOrdersPanel_Click(): RefreshOrdersPanel: End Sub
'�i����PS�{�^�����g���Ȃ�j
'Public Sub UI_RunPipeline_Click():        BtnRunPipeline_Click: End Sub


'----------------------------------------------------------
' �ėp���[�e�B���e�B
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

' ���o�����̂����ŏ��Ɍ���������ԍ���Ԃ��i���������^�C�g�������Ԃ��j
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
' Signals�p�X�F������������΍ŐV��
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
' CSV �捞�iUTF-8/�_�u���N�H�[�g�Ή��j�{ �����~��
'----------------------------------------------------------
Public Sub ImportTopCSV()
    On Error GoTo FAIL

    EnsureDirs
    Dim p As String: p = ResolveSignalsPath()
    If Len(p) = 0 Or Dir$(p) = "" Then
        MsgBox "CSV��������܂���: " & SIGNAL_DIR, vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    ws.Range(START_CELL).CurrentRegion.ClearContents
    ws.Range("A2").value = "Signals: " & p

    ' --- UTF-8 �ǂݍ��� ---
    Dim stm As Object, txt As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile p
    txt = stm.ReadText(-1)
    stm.Close: Set stm = Nothing

    ' --- CSV �� 2�����z��i�_�u���N�H�[�g�Ή��j---
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

    ' --- �V�[�g�֓\�t ---
    ws.Range(START_CELL).Resize(1, cols).value = hdr
    ws.Range(START_CELL).Offset(1, 0).Resize(rCount, cols).value = arr

    ' ts��iB��j�𕶎���Œ�
    ws.Columns(ws.Range(START_CELL).Column + 1).NumberFormatLocal = "@"

    ' �R�[�h��̍��m�iTicker/code/�����R�[�h/Code �̂����ꂩ�j
    Dim hdrRow As Range: Set hdrRow = ws.Range(START_CELL).Resize(1, cols)
    Dim reportCol As String, cT As Long
    cT = FindColAny(hdrRow, reportCol, "Ticker", "code", "�����R�[�h", "Code")
    If cT = 0 Then
        LogEvent "CSV", "NoCodeColumn", "top_candidates.csv �ɃR�[�h�񂪂���܂���iTicker/code/�����R�[�h/Code�j"
    Else
        LogEvent "CSV", "CodeColumn", "Use=" & reportCol
    End If
'================= �Z�N�V���������FImportTopCSV �̏I�[�i�u���j =================
    ' �� J/dJ/vEMA �̐�����~�݁iCSV�捞�̓x�ɍŐV�s�܂Œ��蒼���j
    InstallJKVFormulas False

    ' �� ��荞�ݒ����OHLC��C�i���Y���EHL�����j
    SanitizeOHLC_Repair

    LogEvent "CSV", "ImportTopCSV", p
    SetStatus "CSV�捞����"
    Exit Sub
' �ȍ~ FAIL: �͌��s�ǂ���
'================= �Z�N�V�����I���FImportTopCSV �I�[ =================

FAIL:
    LogEvent "ERR", "ImportTopCSV", Err.Description
    SetStatus "CSV�捞�G���["
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

'================= �Z�N�V�����FSanitizeOHLC_Repair�iCSV��OHLC��C�j =================
' �ړI�F
'   �EHigh < max(Open,Close) / Low > min(Open,Close) ���C
'   �E���Y���i�~10, �~100�j�̎����␳�iVWAP/Close �Ƃ̔�Ŕ���j
'   �EVWAP�������� Close �ő�p
' �Ăяo���F
'   ImportTopCSV �̍Ō�� SanitizeOHLC_Repair ���Ă�
'================= �Z�N�V�����FSanitizeOHLC_Repair�i���O�Փˉ����Łj =================

'================= �Z�N�V�����FSanitizeOHLC_Repair�i�Փˉ���EASCII�j =================
Private Sub SanitizeOHLC_Repair()
    Dim ws As Worksheet, rg As Range, hd As Range
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Set rg = ws.Range(START_CELL).CurrentRegion
    If rg Is Nothing Or rg.rows.Count < 2 Then Exit Sub
    Set hd = rg.rows(1)

    ' ��C���f�b�N�X�icol*�j
    Dim colO As Long, colH As Long, colL As Long, colC As Long, colVW As Long
    colO = FindCol(hd, "Open")
    colH = FindCol(hd, "High")
    colL = FindCol(hd, "Low")
    colC = FindCol(hd, "Close")
    colVW = FindCol(hd, "VWAPd"): If colVW = 0 Then colVW = FindCol(hd, "VWAP")
    If colC = 0 Then Exit Sub

    ' �s�͈�
    Dim r1 As Long, r2 As Long, r As Long
    r1 = rg.row + 1
    r2 = ws.Cells(ws.rows.Count, colC).End(xlUp).row
    If r2 < r1 Then Exit Sub

    ' �l�iv*�j
    Dim vO As Double, vH As Double, vL As Double, vC As Double, vVW As Double
    Dim mx As Double, mn As Double, rat As Double, sc As Double, t As Double

    For r = r1 To r2
        vO = IIf(colO > 0, d(ws.Cells(r, colO).value), 0#)
        vH = IIf(colH > 0, d(ws.Cells(r, colH).value), 0#)
        vL = IIf(colL > 0, d(ws.Cells(r, colL).value), 0#)
        vC = d(ws.Cells(r, colC).value)
        vVW = IIf(colVW > 0, d(ws.Cells(r, colVW).value), 0#)
        If vVW <= 0# Then vVW = vC

        ' 1) ���Y���␳�i�~10/�~100�j
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

        ' 2) HL����
        If colO = 0 Then vO = vC
        If colH = 0 Then vH = vC
        If colL = 0 Then vL = vC
        mx = WorksheetFunction.Max(vO, vC)
        mn = WorksheetFunction.Min(vO, vC)
        If vH < mx Then vH = mx * 1.0001
        If vL > mn Then vL = mn * 0.9999
        If vH < vL Then t = vH: vH = vL: vL = t

        ' 3) ���߂�
        If colO > 0 Then ws.Cells(r, colO).Value2 = vO
        If colH > 0 Then ws.Cells(r, colH).Value2 = vH
        If colL > 0 Then ws.Cells(r, colL).Value2 = vL
        ws.Cells(r, colC).Value2 = vC
        If colVW > 0 And vVW > 0 Then ws.Cells(r, colVW).Value2 = vVW
    Next r

    LogEvent "CSV", "SanitizeOHLC", "rows=" & (r2 - r1 + 1)
End Sub
'================= �Z�N�V�����I���FSanitizeOHLC_Repair =================


'----------------------------------------------------------
' RssMarket �擾�i�w�K�ς݁��������ی��j
'----------------------------------------------------------
Private Function RSS_Get(ByVal code As String, ByVal want As String, Optional ByRef outName As String = "") As Double
    Dim fnames As Variant, fields As Variant
    Dim i As Long, j As Long, v As Variant
    fnames = Split("RssMarket,RssMarket_v,RssMarketEx", ",")

    ' �w�K�ςݗD��
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

    ' �������x��
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

    ' 3) �ی����x���i���`�\�L�𑍓�����j
Dim priceCands As String, vwapCands As String
priceCands = PRICE_LABEL & ",���ݒl(�C�z),���l,���ߒl,Last,LastPrice"
vwapCands = VWAP_LABEL & ",�o�������d����(���i),�o�������d���ρi���i�j,VWAP,���d����,���d����(���i)"

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

            ' �w�K�ۑ��i�u���b�NIf�ň��S�Ɂj
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

' �N�����F�擪�����Ń��x����������
Public Sub ResolveRssLabels(Optional ByVal sampleCode As String = "")
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:   Set rng = ws.Range(START_CELL).CurrentRegion
    If Len(sampleCode) = 0 And rng.rows.Count >= 2 Then
        Dim tit As String, cT As Long
        cT = FindColAny(rng.rows(1), tit, "Ticker", "code", "�����R�[�h", "Code")
        If cT > 0 Then
            sampleCode = TickerToCode(CStr(rng.Cells(2, cT).value))
        Else
            LogEvent "RESOLVE", "NoCodeColumn", "Header not found (Ticker/code/�����R�[�h/Code)"
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
' Live�i3�b�j�FJ/dJ/vEMA ���X�V
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
        ' �񌟏o�iTicker ���̗h��Ή��j
        cT = FindColAny(h, titCode, "Ticker", "code", "�����R�[�h", "Code")
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

            ' �l�̎擾�i�������x���D�恨���t�H�[���o�b�N�j
            px = RSS_Get(code, "PRICE_CAND", nPx)
            vw = RSS_Get(code, "VWAP_CAND", nVw)
            atr = d(rng.Cells(r, cA).value)
            If px <= 0# Or vw <= 0# Or atr <= 0# Then GoTo NXT

            ' ��J
            jp = d(rng.Cells(r, cJ).value)

            ' === �����݃Z�N�V�����i��������j ===
            jn = (px - vw) / atr
            dj = jn - jp
            ve = 0.3 * dj + 0.7 * d(IIf(cvEMA > 0, rng.Cells(r, cvEMA).value, 0))
' --- LiveTick �����ݕ��i��������u���j ---
' Close / VWAPd �͒l�ōX�V
rng.Cells(r, cC).Value2 = px
If cVW > 0 Then rng.Cells(r, cVW).Value2 = vw

' J�񂪐����Ȃ�A��U�l�����Ă���㏑���i�����Z���ł��K������j
If rng.Cells(r, cJ).HasFormula Then rng.Cells(r, cJ).Value2 = rng.Cells(r, cJ).Value2
' J �� dJ �� vEMA �̏��ŋ����㏑���iValue2�j
rng.Cells(r, cJ).Value2 = jn
If cdJ > 0 Then rng.Cells(r, cdJ).Value2 = dj
If cvEMA > 0 Then rng.Cells(r, cvEMA).Value2 = ve
' --- LiveTick �����ݕ��i�u�������܂Łj ---


NXT:    Next r
    End If

RESCHEDULE:
    LiveBusy = False
    Application.OnTime Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC), "LiveTick"
End Sub


' �����N���b�N���؁i���O�ł�J/dJ/vEMA�̍X�V���m�F�j
Public Sub LiveOnce_Debug()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:   Set rng = ws.Range(START_CELL).CurrentRegion
    If rng.rows.Count < 2 Then MsgBox "��₪����܂���": Exit Sub

    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cvEMA&, cVW&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "�����R�[�h", "Code")
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
                  " J:" & Format$(jp, "0.###") & "��" & Format$(jn, "0.###")
            cnt = cnt + 1
        Else
            rep = rep & IIf(Len(rep) > 0, " | ", "") & code & ":px/vw/atr=0"
        End If
NXT:
    Next r
    LogEvent "LIVEDBG", "LiveOnce", rep
    MsgBox "LiveOnce_Debug �����FLOG �� LIVEDBG | LiveOnce ���m�F���Ă��������B", vbInformation
End Sub

'----------------------------------------------------------
' �f�������FCSV�ۑ��{Orders�ɑ��ǋL
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
    If Trim$(ws.Range("A1").value) <> "�� �����ꗗ" Then
        ws.Cells.Clear
        ws.Range("A1").value = "�� �����ꗗ": ws.Range("A1").Font.Bold = True
        ws.Range("A4").value = "�� ���ꗗ": ws.Range("A4").Font.Bold = True
        ws.Range("A2").value = "(�f���������O������܂���)"
    End If
    Dim row As Long: row = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If row < 2 Then row = 1
    If ws.Range("A2").value = "(�f���������O������܂���)" Then ws.Range("A2").ClearContents: row = 1
    With ws
        .Cells(row + 1, 1).value = Format$(Now, "yyyy-MM-dd HH:mm:ss")
        .Cells(row + 1, 2).value = "DEMO-ORDER"
        .Cells(row + 1, 3).value = side & " " & code & " x" & qty & " @" & Format$(price, "0.000")
        .Cells(row + 1, 4).value = status
    End With
End Sub

'----------------------------------------------------------
' Orders�p�l���i�K�v�ɉ����Ė{�Ԃ�RSS�ꗗ�ɍ��։j
'----------------------------------------------------------
Public Sub RefreshOrdersPanel()
    Dim ws As Worksheet: On Error Resume Next: Set ws = ThisWorkbook.Worksheets(ORD_SHEET): On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.name = ORD_SHEET
    ws.Columns.AutoFit
End Sub

'----------------------------------------------------------
' BookNew�F�t�����Z�b�g�{�{�^���Ċ����{J/dJ/vEMA �̐����~��
'----------------------------------------------------------
Public Sub BookNew()
    On Error Resume Next

    '---- 1) ���S��~�i�^�C�}�[�����j
    Application.OnTime Now + TimeSerial(0, 0, 1), "LiveTick", , False
    LiveBusy = False: CsvBusy = False

    '---- 2) �t�H���_�m��
    EnsureDirs

    '---- 3) Dashboard ���Đ���
    Dim wsD As Worksheet
    If SheetExists(DASH_SHEET) Then
        Set wsD = ThisWorkbook.Worksheets(DASH_SHEET)
        wsD.Cells.Clear
    Else
        Set wsD = ThisWorkbook.Worksheets.Add
        wsD.name = DASH_SHEET
    End If

    With wsD
        .Range("A1").value = "�t����X�L�����s���O Dashboard�iBookNew�j"
        .Range("A2").value = "Signals: " & ResolveSignalsPath()
       .Range(START_CELL).Resize(1, 24).value = Array( _
    "ts", "Open", "High", "Low", "Close", "Volume", "VWAPd", "ATR5", "J", "dJ", "d2J", "vEMA", _
    "Session", "Ticker", "IBS", "SMA20", "STD20", "Z20", "ROC5", "Turnover", "p", "side", "TP_price", "SL_price", "Value")
        .Columns.ColumnWidth = 12
    End With

    '---- 4) LOG �V�[�g
    Dim wsL As Worksheet
    If SheetExists(LOG_SHEET) Then
        Set wsL = ThisWorkbook.Worksheets(LOG_SHEET): wsL.Cells.Clear
    Else
        Set wsL = ThisWorkbook.Worksheets.Add: wsL.name = LOG_SHEET
    End If
    wsL.Range("A1:D1").value = Array("Time", "Kind", "Detail", "Note")
    wsL.Range("A1:D1").Font.Bold = True
    wsL.Columns("A:D").ColumnWidth = 26

    '---- 5) Orders �V�[�g
    Dim wsO As Worksheet
    If SheetExists(ORD_SHEET) Then
        Set wsO = ThisWorkbook.Worksheets(ORD_SHEET): wsO.Cells.Clear
    Else
        Set wsO = ThisWorkbook.Worksheets.Add: wsO.name = ORD_SHEET
    End If
    With wsO
        .Range("A1").value = "�� �����ꗗ": .Range("A1").Font.Bold = True
        .Range("A4").value = "�� ���ꗗ": .Range("A4").Font.Bold = True
        .Range("A2").value = "(�f���������O������܂���)"
        .Columns.AutoFit
    End With

    '---- 6) Dashboard �{�^���iUI_* �n���h���ɓ���j
    AddOrResetBtn wsD, "btnUpdateAll", "�����X�V�iPS��CSV�j", "UI_UpdateAll_Click", "C2", 70, 20
  AddOrResetBtn wsD, "btnImport", "�X�V�iCSV�捞�j", "UI_Import_Click", "I2", 70, 20
AddOrResetBtn wsD, "btnAutoStartDemo", "�����J�n�i�f���j", "UI_AutoStartDemo_Click", "D2", 70, 20
AddOrResetBtn wsD, "btnAutoStartLive", "�����J�n�i�{�ԁj", "UI_AutoStartLive_Click", "E2", 70, 20
AddOrResetBtn wsD, "btnAutoStop", "��~�i�p�j�b�N�j", "UI_AutoStop_Click", "F2", 70, 20
AddOrResetBtn wsD, "btnOrders", "�����󋵃p�l���X�V", "UI_RefreshOrdersPanel_Click", "H2", 70, 20
'�i����PS�{�^�����g���ꍇ�j
AddOrResetBtn wsD, "btnPS", "����PS ���s", "UI_RunPipeline_Click", "J2", 70, 20

AddOrResetBtn wsD, "btnSimTradeStart", "���s�J�n", "UI_SimTradeStart_Click", "K2", 70, 20
AddOrResetBtn wsD, "btnSimTradeStop", "���s��~", "UI_SimTradeStop_Click", "L2", 70, 20
AddOrResetBtn wsD, "btnSimTradeClear", "���s�N���A", "UI_SimTradeClear_Click", "M2", 70, 20


    '---- 7) J/dJ/vEMA �̐�����~�݁i�ی�s�v�Ȃ� False�A�㏑���s�ɂ���Ȃ� True�j
    InstallJKVFormulas False

    LogEvent "SYS", "SetupNewBook", "ready"
    SetStatus "ready"
    MsgBox "BookNew �����FDashboard/LOG/Orders ���Đ������AJ/dJ/vEMA �̐�����~�݂��܂����B", vbInformation
End Sub

' Dashboard�Ƀ{�^����1�쐬�i����������΍폜���č쐬�j
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
' �����J�n/��~�i�ŏ��j
'----------------------------------------------------------
Public Sub AutoStart(ByVal demo As Boolean)
EnsureLogInfra:     EnsureDirs
    DEMO_MODE = demo: g_OrderID = 1
    ImportTopCSV
    ResolveRssLabels
    LiveStart
    LogEvent "SYS", "AutoStart", IIf(DEMO_MODE, "DEMO", "LIVE"): SetStatus IIf(demo, "�����J�n�i�f���j", "�����J�n�i�{�ԁj")
End Sub

Public Sub AutoStop()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC), "LiveTick", , False
    LogEvent "SYS", "LiveStop", "": LogEvent "SYS", "AutoStop", "": SetStatus "��~�i�p�j�b�N�j"
End Sub

'----------------------------------------------------------
' Pipeline�i�񓯊��j���K�v�ȂƂ������g�p
'----------------------------------------------------------
Public Sub BtnRunPipeline_Click()
    On Error GoTo FIN
    EnsureDirs
    Dim ps As String, script As String, log As String, wrap As String
    ps = GetPSPath(): If Len(ps) = 0 Then MsgBox "PowerShell ��������܂���B", vbExclamation: Exit Sub
    script = BASE_DIR & "\scripts\run_full.ps1"
    If Dir$(script) = "" Then MsgBox "run_full.ps1 ��������܂���: " & script, vbExclamation: Exit Sub
    log = LOG_DIR & "\pipeline_" & Format(Now, "yyyy-MM-dd_HHnnss") & ".txt"
    wrap = LOG_DIR & "\run_wrapper_" & Format(Now, "yyyy-MM-dd_HHnnss") & ".ps1"
    If Not WriteWrapper(script, log, wrap) Then MsgBox "PowerShell ���b�p�[�쐬�Ɏ��s", vbExclamation: Exit Sub
    RunPS_Async ps, wrap
    LogEvent "PS", "Pipeline launch", log
    SetStatus "�N��: " & log
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

'========================  �x����V�~�����[�^�[  ========================
' Close/VWAPd �������_���E�H�[�N�����AJ/dJ/vEMA ��A���X�V���܂��B
' �g�����F
'   Alt+F8 �� SimStart_Fast�i1�b/�s50/�{��0.10%�j or  SimStart_Slow�i3�b/�s200/�{��0.05%�j
'   Alt+F8 �� SimStop�i��~�j �� SimReset�i��荞�ݒ���ɖ߂��j
'======================================================================


'================= �Z�N�V�����ۂ��ƁFSimStart_Fast/Slow/SimStart =================
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

    ' �d�v�F�p�X��Ԃ��������i�O��̒�~�ʒu����������Ȃ��j
    If Not gSimPath Is Nothing Then gSimPath.RemoveAll

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then
        MsgBox "��₪����܂���B��� [�X�V�iCSV�捞�j]�B", vbExclamation
        Exit Sub
    End If

    Dim h As Range: Set h = rng.rows(1)
    Dim cC&, cVW&: cC = FindCol(h, "Close"): cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    Dim lastRow&, useN&, r&
    lastRow = ws.Cells(ws.rows.Count, cC).End(xlUp).row
    useN = Application.Min(rows, lastRow - (rng.row + 1) + 1)
    If useN <= 0 Then MsgBox "�f�[�^�s������܂���B", vbExclamation: Exit Sub

    ReDim simSnap(1 To useN, 1 To 2)
    For r = 1 To useN
        simSnap(r, 1) = d(ws.Cells(rng.row + r, cC).value)
        simSnap(r, 2) = d(IIf(cVW > 0, ws.Cells(rng.row + r, cVW).value, ws.Cells(rng.row + r, cC).value))
    Next r

    simRunning = True
    LogEvent "SIM", "Start", "vol=" & simVolPct & " step=" & simStepSec & "s rows=" & useN
    SimTick     ' ���񑦎�
End Sub
'================= �Z�N�V�����I���FSimStart =================





'================= �Z�N�V�����ۂ��ƁFSimStop =================
Public Sub SimStop()
    CancelSim
    simRunning = False
    LogEvent "SIM", "Stop", ""
End Sub
'================= �Z�N�V�����I���FSimStop =================




Public Sub SimReset()
    If simRunning Then
        MsgBox "SimStop �̌�Ɏ��s���Ă��������B", vbInformation
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
'================= �Z�N�V�����I���FSimStart / Stop / Reset =================



'=================  �Z�N�V�����ۂ��ƍ����ւ��FSimTick  =================
' �ESimStart_* ����X�P�W���[������A�x����ł� Close/VWAPd ���[���I�ɍX�V
' �EJ/dJ/vEMA �́AJ�񂪐����Ȃ�V�[�g�v�Z�ɁA���l�Ȃ�Value2�ŏ㏑��
' �E���i�X�V��� SimTrade_OnTick ���Ă�ŁAENTRY/EXIT�`Orders�L�^�܂ōs��
'======================================================================
'=================  �Z�N�V�����ۂ��ƁFSimTick�iOHLC���炩�p�X�j  =================
'================= �Z�N�V�����ۂ��ƁFSimTick�i�f�f�{OHLC���炩�p�X�j =================
'================= �Z�N�V�����ۂ��ƁFSimTick�i�R�[�h��Ȃ��ϐ��Łj =================
Public Sub SimTick()
    On Error GoTo SCHEDULE
    If Not simRunning Then GoTo SCHEDULE

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then GoTo SCHEDULE

    ' �w�b�_
    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cOpen&, cHigh&, cLow&, cClose&, cVW&, cA&, cJ&, cdJ&, cvEMA&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "�����R�[�h", "Code", "�R�[�h")
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

    ' �͈�
    Dim firstRow As Long, lastRow As Long, useN As Long
    firstRow = rng.row + 1
    lastRow = ws.Cells(ws.rows.Count, cClose).End(xlUp).row
    If lastRow < firstRow Then GoTo SCHEDULE
    useN = Application.Min(simRows, lastRow - firstRow + 1)

    ' ���i�X�V�iOHLC���炩�p�X�j
    Dim r As Long, key$, px#, vw#, atr#, jp#, jn#, dj#, ve#
    For r = firstRow To firstRow + useN - 1
        ' �R�[�h�񂪖����ꍇ�͍s�ԍ��ŋ[���L�[
        If cT > 0 Then
            key = TickerToCode(CStr(ws.Cells(r, cT).value))
            If Len(key) <> 4 Then key = "R" & CStr(r)
        Else
            key = "R" & CStr(r)
        End If

        px = SimPath_NextPrice_OHLC(ws, r, cOpen, cHigh, cLow, cClose, cVW, key, vw)

        ' ���߂�
        ws.Cells(r, cClose).Value2 = px
        If cVW > 0 Then ws.Cells(r, cVW).Value2 = vw

        ' J/dJ/vEMA�i�����Z���͌v�Z�ɔC����j
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
'================= �Z�N�V�����I���FSimTick =================





' �W�����K�����iBox-Muller�j
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


'=== �I�𒆂̃f�[�^�s�� J/dJ/vEMA �� RSS �l�ŋ����X�V ===
Public Sub ForceUpdateJ_ActiveRow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then MsgBox "��₪����܂���": Exit Sub

    Dim r As Long: r = ActiveCell.row
    If r <= rng.row Then r = rng.row + 1      ' ���o���s�̂Ƃ��� 1 �s��

    ForceUpdateJ_Core r
End Sub

'=== �擪�f�[�^�s�iB5�� 1 �s���j�������X�V ===
Public Sub ForceUpdateJ_FirstDataRow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then MsgBox "��₪����܂���": Exit Sub
    ForceUpdateJ_Core rng.row + 1
End Sub

'=== ���́F�s�ԍ� r �� J/dJ/vEMA �� RSS�����ōX�V�iLOG�ɏڍׂ��o�́j ===
Private Sub ForceUpdateJ_Core(ByVal r As Long)
    On Error GoTo FAIL

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then
        LogEvent "FORCEJ", "NoData", "rng is empty": Exit Sub
    End If

    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cvEMA&, cVW&, tit As String
    cT = FindColAny(h, tit, "Ticker", "code", "�����R�[�h", "Code")
    cC = FindCol(h, "Close"): cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J")
    cdJ = FindCol(h, "dJ"):   cvEMA = FindCol(h, "vEMA")
    cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")

    If cC * cA * cJ = 0 Then
        LogEvent "FORCEJ", "HeaderMissing", "cT=" & cT & "(" & tit & "), cC=" & cC & ", cA=" & cA & ", cJ=" & cJ
        Exit Sub
    End If

    ' �͈̓`�F�b�N
    If r < rng.row + 1 Or r > rng.row + rng.rows.Count - 1 Then
        LogEvent "FORCEJ", "RowOut", "r=" & r: Exit Sub
    End If

    ' �R�[�h�擾�i��ɕ����񉻁������R�[�h���j
    Dim tkr As String, code As String
    tkr = Trim$(CStr(rng.Cells(r, cT).value))
    code = TickerToCode(tkr)
    If Len(code) <> 4 Then
        LogEvent "FORCEJ", "BadCode", "tkr=" & tkr & " code=" & code
        Exit Sub
    End If

    ' RSS �擾�iVariant�����S���l���j
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
        MsgBox "px/vw/atr �̂����ꂩ�� 0 �ōX�V�ł��܂���B", vbExclamation
        Exit Sub
    End If

    ' J/dJ/vEMA �v�Z��������
    Dim jp As Double, jn As Double, dj As Double, ve As Double
    jp = d(rng.Cells(r, cJ).value)
    jn = (px - vw) / atr
    dj = jn - jp
    ve = 0.3 * dj + 0.7 * d(IIf(cvEMA > 0, rng.Cells(r, cvEMA).value, 0))

 ' --- FORCEJ �����ݕ��i��������u���j ---
rng.Cells(r, cC).Value2 = px
If cVW > 0 Then rng.Cells(r, cVW).Value2 = vw

If rng.Cells(r, cJ).HasFormula Then rng.Cells(r, cJ).Value2 = rng.Cells(r, cJ).Value2
rng.Cells(r, cJ).Value2 = jn
If cdJ > 0 Then rng.Cells(r, cdJ).Value2 = dj
If cvEMA > 0 Then rng.Cells(r, cvEMA).Value2 = ve
' --- FORCEJ �����ݕ��i�u�������܂Łj ---


    LogEvent "FORCEJ", "Jupdated", code & " " & Format$(jp, "0.###") & "��" & Format$(jn, "0.###")
    Exit Sub
FAIL:
    LogEvent "ERR", "FORCEJ", Err.Description
    MsgBox "FORCEJ �ŃG���[: " & Err.Description, vbExclamation
End Sub




'=== Dashboard �́u����PS ���s�v�{�^���d���̉������Ĕz�u ===
Public Sub FixPipelineButton()
    Dim ws As Worksheet, shp As Shape, hit As Long
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)

    ' 1) �����́u����PS ���s�v���ۂ��{�^����S�������i�L���v�V����/OnAction���ʂŔ���j
    For Each shp In ws.Shapes
        On Error Resume Next
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                Dim cap$: cap = ""
                cap = shp.TextFrame.Characters.Text
                If cap Like "*����PS*" Or shp.OnAction = "UI_RunPipeline_Click" Or shp.OnAction = "BtnRunPipeline_Click" Then
                    shp.Delete
                End If
            End If
        End If
        On Error GoTo 0
    Next shp

    ' 2) ���K��PS�{�^����1�����z�u�i�ʒu�� J2�F�K�v�Ȃ�ς���OK�j

AddOrResetBtn ws, "btnPS", "����PS ���s", "UI_RunPipeline_Click", "J2", 120, 28


    LogEvent "FIX", "PipelineButton", "Reinstalled"
End Sub

'=== J / dJ / d2J / vEMA �̐�����~�݁i�S�s�E����j ===
'================= �Z�N�V�����ۂ��ƍ����ւ��FInstallJKVFormulas�i�V�[�h�Łj =================
Public Sub InstallJKVFormulas(Optional ByVal protectCols As Boolean = False)
    On Error GoTo FIN

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then Exit Sub

    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cD2&, cvEMA&, cVW&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "�����R�[�h", "Code")
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

    ' R1C1 �I�t�Z�b�g
    Dim offC&, offVW&, offATR&: offC = cC - cJ: offVW = cVW - cJ: offATR = cA - cJ

    ' J: (Close - VWAPd) / ATR5
    Dim fJ$: fJ = "=(RC[" & offC & "]-RC[" & offVW & "])/RC[" & offATR & "]"
    ws.Range(ws.Cells(firstRow, cJ), ws.Cells(lastRow, cJ)).FormulaR1C1 = fJ

    ' dJ: �擪= Jr - J(r+1) / �ȍ~=����
    If cdJ > 0 Then
        Dim offJ_from_dJ&: offJ_from_dJ = cJ - cdJ
        ws.Cells(firstRow, cdJ).FormulaR1C1 = "=RC[" & offJ_from_dJ & "]-R[1]C[" & offJ_from_dJ & "]"
        If firstRow + 1 <= lastRow Then
            ws.Range(ws.Cells(firstRow + 1, cdJ), ws.Cells(lastRow, cdJ)).FormulaR1C1 = _
                "=RC[-" & (cdJ - cJ) & "]-R[-1]C"
        End If
    End If

    ' d2J: �擪= dJr - dJ(r+1) / �ȍ~=����
    If cD2 > 0 And cdJ > 0 Then
        Dim off_dJ_from_d2&: off_dJ_from_d2 = cdJ - cD2
        ws.Cells(firstRow, cD2).FormulaR1C1 = "=RC[" & off_dJ_from_d2 & "]-R[1]C[" & off_dJ_from_d2 & "]"
        If firstRow + 1 <= lastRow Then
            ws.Range(ws.Cells(firstRow + 1, cD2), ws.Cells(lastRow, cD2)).FormulaR1C1 = _
                "=RC[-" & (cD2 - cdJ) & "]-R[-1]C"
        End If
    End If

    ' vEMA: �擪= dJ�iK��j/ �ȍ~= 0.3*dJ + 0.7*�O��
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
'=================  �Z�N�V�����ۂ��ƍ����ւ��FInstallJKVFormulas�i�V�[�h�Łj�����܂�  =================



'================= �Z�N�V�����ۂ��ƁFUI�i�{�^���j =================
Public Sub UI_SimTradeStart_Click():  SimTradeStart_Engine:  End Sub
Public Sub UI_SimTradeStop_Click():   SimTradeStop_Engine:   End Sub
Public Sub UI_SimTradeClear_Click():  SimTradeClear_Engine:  End Sub
'================= �Z�N�V�����I���FUI =================

'============= �Z�N�V�����ۂ��ƁF�^�����s Start/Stop/Clear(Engine) =============
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
'============= �Z�N�V�����I���FStart/Stop/Clear(Engine) =============

'================= ��������̓R�����g�̂݁i����j ==================
' �E���̉��ɂ͎��s�R�[�h��u���Ȃ����ƁiEnd Sub�̌�̓R�����g�̂݋��j
' �ESimTick�̖����� SimTrade_OnTick ���Ăԑz��
' �EOrders �ւ̋L�^�� SimLogEntry / SimLogExit ���g�p
'===================================================================


'=================  �Z�N�V�����ۂ��ƍ����ւ��FSimTrade_OnTick �����܂�  =================

'=================  �Z�N�V�����ۂ��ƍ����ւ��FSimTrade_OnTick �����܂�  =================


'==== Orders �ւ̋L�^�iENTRY�j6������ ====
Private Sub SimLogEntry(ByVal code As String, ByVal side As String, _
                        ByVal qty As Long, ByVal px As Double, _
                        ByVal tp As Double, ByVal sl As Double)

    Dim ws As Worksheet
    If SheetExists(ORD_SHEET) = False Then
        Set ws = ThisWorkbook.Worksheets.Add: ws.name = ORD_SHEET
    Else
        Set ws = ThisWorkbook.Worksheets(ORD_SHEET)
    End If

    ' �w�b�_������
    If Trim$(ws.Range("A1").value) <> "�� �����ꗗ" Then
        ws.Cells.Clear
        ws.Range("A1").value = "�� �����ꗗ": ws.Range("A1").Font.Bold = True
        ws.Range("A4").value = "�� ���ꗗ": ws.Range("A4").Font.Bold = True
        ws.Range("A2").value = "(�f���������O������܂���)"
    End If

    ' �s�ǉ�
    Dim row As Long: row = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If row < 2 Then row = 1
    If ws.Range("A2").value = "(�f���������O������܂���)" Then ws.Range("A2").ClearContents: row = 1

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


'=================  �Z�N�V�����ǉ��FSimTrade_OnTick + �N��/��~  =================
' �ݒ�l�i���錾�Ȃ�擪�ɐ錾�j
'Private Const SIM_TH_ABSJ   As Double = 0.8
'Private Const SIM_K_TP_ATR  As Double = 2#
'Private Const SIM_K_SL_ATR  As Double = 2#
'Private Const SIM_TMAX_BAR  As Long   = 60
'Private Const SIM_MIN_REENTRY As Long = 20
'Private Const SIM_QTY       As Long   = 100
'Private simTradeOn As Boolean
'Private gSimPos As Object   ' Scripting.Dictionary



'=================  �Z�N�V�����ۂ��ƁFSimTrade_OnTick�i���S�Łj  =================
' �ESimTick �� Close/VWAPd ���X�V�����g��h�ɌĂԁiSimTick���� SCHEDULE: �ŌĂԁj
' �EEXIT �� ENTRY �̏��ŏ������AOrders �� SIM-EXIT / SIM-ENTRY ��ǋL
' �ELOG �� Digest ���o���A�������o�Ȃ�����������iprice0/atrAdj/hold/cool/noSignal/badCode�j
Public Sub SimTrade_OnTick()

    If Not simTradeOn Then Exit Sub
    On Error GoTo FIN

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Or rng.rows.Count < 2 Then Exit Sub

    '--- �w�b�_���o�i�\�L��炬�Ή��j ---
    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cC&, cA&, cJ&, cdJ&, cD2&, cvEMA&, cVW&, tit As String
    cT = FindColAny(h, tit, "Ticker", "code", "�����R�[�h", "Code")
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

    '--- �����͈� ---
    Dim firstRow As Long, lastRow As Long
    firstRow = rng.row + 1
    lastRow = ws.Cells(ws.rows.Count, cC).End(xlUp).row
    If lastRow < firstRow Then Exit Sub

    If gSimPos Is Nothing Then Set gSimPos = CreateObject("Scripting.Dictionary")

    '============================== 1) EXIT �� ==============================
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

                ' ���Ԑ؂�
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

        ' ���S�N���[�Y�iopen=False ���� cool<=0�j������
        Dim del As New Collection
        For Each k In gSimPos.Keys
            Set pos = gSimPos(k)
            If (Not pos("open")) And pos("cool") <= 0 Then del.Add k
        Next k
        For Each k In del: gSimPos.Remove k: Next k
    End If

    '============================== 2) ENTRY ==============================
    Dim row As Long, code$, j#, dj#, ve#, side$, tp#, sl#, qty&, vw#
    Dim px0&, atr0&, sig0&, hold0&, cd0&  ' digest �p�J�E���^
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

        ' �ۗL/�N�[���_�E�����̓X�L�b�v
        If gSimPos.Exists(code) Then
            Set pos = gSimPos(code)
            If pos("open") Or pos("cool") > 0 Then hold0 = hold0 + 1: GoTo NextRow
        End If

        ' �t����V�O�i��
        If Abs(j) >= SIM_TH_ABSJ And Sgn(ve) <> Sgn(dj) Then
            side = IIf(j < 0, "BUY", "SELL")
            If side = "BUY" Then
                tp = px + SIM_K_TP_ATR * atr: sl = px - SIM_K_SL_ATR * atr
            Else
                tp = px - SIM_K_TP_ATR * atr: sl = px + SIM_K_SL_ATR * atr
            End If

            ' �|�W�V�����ۑ�
            Set pos = CreateObject("Scripting.Dictionary")
            pos("open") = True: pos("code") = code: pos("side") = side
            pos("qty") = qty: pos("px") = px: pos("tp") = tp: pos("sl") = sl
            pos("row") = row: pos("bars") = 0: pos("cool") = 0
            gSimPos(code) = pos

            ' Orders�֋L�^
            SimLogEntry code, side, qty, px, tp, sl
        Else
            sig0 = sig0 + 1
        End If

NextRow:
    Next row

    ' �_�C�W�F�X�g�i�������o�Ȃ��Ƃ��̓���j
    LogEvent "SIMTRADE", "Digest", _
        "Rows=" & (lastRow - firstRow + 1) & _
        " price0=" & px0 & " atrAdj=" & atr0 & " hold/cool=" & hold0 & " noSignal=" & sig0 & " badCode=" & cd0

FIN:
    Exit Sub
End Sub
'=================  �Z�N�V�����ۂ��ƁFSimTrade_OnTick�i���S�Łj�����܂�  =================


'=================  �Z�N�V�����ǉ��FSimTrade_OnTick �����܂�  =================

'================= �Z�N�V�����ǉ��FUpdateAll�iPS��Import���������x���j =================
Public Sub UI_UpdateAll_Click()
    UpdateAll_Run 300   ' �ő�ҋ@300�b�i���ɍ��킹�Ē����j
End Sub

'================= �Z�N�V�����ۂ��ƁFUpdateAll_Run�iPS���s�ł����s�j =================
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
        ' ���s���Ă������p���iMsgBox�͏o���Ȃ��j
    Else
        LogEvent "PS", "Pipeline skip", "ps/script/Wrapper NG"
    End If

    ImportTopCSV
    InstallJKVFormulas False
    ResolveRssLabels
    SetStatus "UpdateAll ����"
    Exit Sub
FAIL:
    LogEvent "ERR", "UpdateAll", Err.Description
    ' ���s���Ȃ�
End Sub
'================= �Z�N�V�����I���FUpdateAll_Run =================


' �������s�FPS��ҋ@���A�ő� maxWaitSec �b�Ń^�C���A�E�g�B�߂�l�� exit code�i-1=�^�C���A�E�g�j
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
'================= �Z�N�V�����ǉ��FUpdateAll �����܂� =================




'================= �Z�N�V�����ǉ��F�蓮�L�b�N�i�f�f�p�j =================
Public Sub UI_SimKick_Click()
    LogEvent "SIMDBG", "Kick", "manual"
    SimTick
End Sub
'================= �Z�N�V�����ǉ������܂� =================




'================= �Z�N�V�����FDiag_WriteTest =================
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
'================= �Z�N�V�����I���FDiag_WriteTest =================

'================= �Z�N�V�����FSimTick_Minimal =================
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
        ' �����_���E�H�[�N
        px = px * (1 + RandN() * 0.001)
        ws.Cells(r, cC).Value2 = px

        If cVW > 0 Then
            vw = d(ws.Cells(r, cVW).value): If vw <= 0# Then vw = px
            vw = vw + 0.15 * (px - vw)           ' �Ȉ�VWAP�ǐ�
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
'================= �Z�N�V�����I���FSimTick_Minimal =================

'================= �Z�N�V�����FDiag_UnprotectAll =================
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
'================= �I��� =================

'================= �Z�N�V�����FDiag_ReportHeaders =================
Public Sub Diag_ReportHeaders()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim rng As Range:    Set rng = ws.Range(START_CELL).CurrentRegion
    If rng Is Nothing Then LogEvent "DIAG", "NoRange", START_CELL: Exit Sub
    Dim h As Range: Set h = rng.rows(1)
    Dim cT&, cO&, cH&, cL&, cC&, cVW&, cA&, cJ&, cdJ&, cvE&, tit$
    cT = FindColAny(h, tit, "Ticker", "code", "�����R�[�h", "Code", "�R�[�h")
    cO = FindCol(h, "Open"): cH = FindCol(h, "High"): cL = FindCol(h, "Low")
    cC = FindCol(h, "Close"): cVW = FindCol(h, "VWAPd"): If cVW = 0 Then cVW = FindCol(h, "VWAP")
    cA = FindCol(h, "ATR5"): cJ = FindCol(h, "J"): cdJ = FindCol(h, "dJ"): cvE = FindCol(h, "vEMA")
    LogEvent "DIAG", "Headers", _
      "start=" & START_CELL & " firstRow=" & rng.row + 1 & _
      " cT=" & cT & "(" & tit & ")" & " cC=" & cC & " cVW=" & cVW & _
      " cA=" & cA & " cJ=" & cJ & " cdJ=" & cdJ & " cvEMA=" & cvE
End Sub
'================= �I��� =================


'================= �Z�N�V�����FDiag_HardWriteLoop =================
Public Sub Diag_HardWriteLoop()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Dim tgt As Range:    Set tgt = ws.Range("Z2")   ' �ǂ��ł��ǂ��󂫃Z��
    Dim i As Long
    For i = 1 To 10
        tgt.Value2 = "PING " & Format$(Now, "HH:NN:SS")
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
    LogEvent "DIAG", "HardWriteLoop", "done @" & Format$(Now, "HH:NN:SS")
End Sub
'================= �I��� =================

