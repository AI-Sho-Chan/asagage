Attribute VB_Name = "FixHI_Min"
Option Explicit

Private Function DayKey(ByVal v As Variant) As Long
    On Error GoTo fail
    If IsDate(v) Then DayKey = CLng(CDbl(CDate(v))): Exit Function
    If IsNumeric(v) Then DayKey = CLng(CDbl(v)): Exit Function
fail:
    DayKey = 0
End Function

Private Function BlockLastRow(ByVal sc As Long) As Long
    Dim ws As Worksheet
    Dim j As Long
    Dim mx As Long
    Set ws = Worksheets("Bars")
    mx = 3
    For j = sc To sc + 9
        Dim lr As Long
        lr = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
        If lr > mx Then mx = lr
    Next j
    BlockLastRow = mx
End Function

Private Function VWAP_FromBars(ByVal sc As Long) As Variant
    Dim ws As Worksheet
    Set ws = Worksheets("Bars")

    Const offD As Long = 3   'E列=日付
    Const offC As Long = 8   'J列=終値
    Const offV As Long = 9   'K列=出来高

    Dim colD As Long: colD = sc + offD
    Dim colC As Long: colC = sc + offC
    Dim colV As Long: colV = sc + offV

    Dim last As Long: last = BlockLastRow(sc)
    Dim r As Long, k As Long
    Dim haveToday As Boolean
    Dim keyToday As Long: keyToday = CLng(Date)
    Dim vSum As Double, pvSum As Double

    For r = 3 To last
        k = DayKey(ws.Cells(r, colD).Value2)
        If k = keyToday Then
            haveToday = True
            If IsNumeric(ws.Cells(r, colV).Value2) And IsNumeric(ws.Cells(r, colC).Value2) Then
                vSum = vSum + CDbl(ws.Cells(r, colV).Value2)
                pvSum = pvSum + CDbl(ws.Cells(r, colV).Value2) * CDbl(ws.Cells(r, colC).Value2)
            End If
        End If
    Next r
    If haveToday And vSum > 0 Then
        VWAP_FromBars = pvSum / vSum
        Exit Function
    End If

    Dim keyLast As Long: keyLast = 0
    vSum = 0: pvSum = 0
    For r = last To 3 Step -1
        k = DayKey(ws.Cells(r, colD).Value2)
        If k > 0 Then keyLast = k: Exit For
    Next r
    If keyLast = 0 Then VWAP_FromBars = CVErr(xlErrNA): Exit Function

    For r = 3 To last
        If DayKey(ws.Cells(r, colD).Value2) = keyLast Then
            If IsNumeric(ws.Cells(r, colV).Value2) And IsNumeric(ws.Cells(r, colC).Value2) Then
                vSum = vSum + CDbl(ws.Cells(r, colV).Value2)
                pvSum = pvSum + CDbl(ws.Cells(r, colV).Value2) * CDbl(ws.Cells(r, colC).Value2)
            End If
        End If
    Next r

    If vSum > 0 Then VWAP_FromBars = pvSum / vSum Else VWAP_FromBars = CVErr(xlErrNA)
End Function

Private Function ATR5_FromBars(ByVal sc As Long) As Variant
    Dim ws As Worksheet
    Set ws = Worksheets("Bars")

    Const offD As Long = 3   'E=日付
    Const offH As Long = 6   'H=高値
    Const offL As Long = 7   'I=安値
    Const offC As Long = 8   'J=終値

    Dim colD As Long: colD = sc + offD
    Dim colH As Long: colH = sc + offH
    Dim colL As Long: colL = sc + offL
    Dim colC As Long: colC = sc + offC

    Dim last As Long: last = BlockLastRow(sc)
    Dim r As Long, k As Long
    Dim keyToday As Long: keyToday = CLng(Date)
    Dim keyTarget As Long: keyTarget = 0
    Dim haveToday As Boolean

    For r = last To 3 Step -1
        k = DayKey(ws.Cells(r, colD).Value2)
        If k = keyToday Then keyTarget = keyToday: haveToday = True: Exit For
        If keyTarget = 0 And k > 0 Then keyTarget = k
    Next r
    If keyTarget = 0 Then ATR5_FromBars = CVErr(xlErrNA): Exit Function

    Dim cnt As Long: cnt = 0
    Dim trSum As Double: trSum = 0
    Dim prevC As Double: prevC = 0

    For r = last To 3 Step -1
        If DayKey(ws.Cells(r, colD).Value2) <> keyTarget Then
            If haveToday Then Exit For
            GoTo SkipRow
        End If
        If IsNumeric(ws.Cells(r, colH).Value2) And IsNumeric(ws.Cells(r, colL).Value2) And IsNumeric(ws.Cells(r, colC).Value2) Then
            Dim hi As Double: hi = CDbl(ws.Cells(r, colH).Value2)
            Dim lo As Double: lo = CDbl(ws.Cells(r, colL).Value2)
            Dim cls As Double: cls = CDbl(ws.Cells(r, colC).Value2)
            Dim tr As Double
            If cnt = 0 Then prevC = cls
            tr = Application.Max(hi - lo, Abs(hi - prevC), Abs(lo - prevC))
            trSum = trSum + tr
            cnt = cnt + 1
            prevC = cls
            If cnt = 5 Then Exit For
        End If
SkipRow:
    Next r

    If cnt > 0 Then ATR5_FromBars = trSum / cnt Else ATR5_FromBars = CVErr(xlErrNA)
End Function

Public Sub FixHI_Run()
    Dim D As Worksheet
    Set D = Worksheets("Dashboard")
    Dim last As Long: last = D.Cells(D.Rows.Count, "A").End(xlUp).Row

    Dim r As Long, code As String, sc As Long
    Dim vH As Variant, vI As Variant

    Application.ScreenUpdating = False
    For r = 2 To last
        code = CStr(D.Cells(r, "A").Value)
        If Len(code) = 4 Then
            sc = M_ASG_Core.FindBlockCol(code)
            If sc > 0 Then
                vH = VWAP_FromBars(sc)
                vI = ATR5_FromBars(sc)
                If Not IsError(vH) Then D.Cells(r, "H").Value = vH
                If Not IsError(vI) Then D.Cells(r, "I").Value = vI
            End If
        End If
    Next r
    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "H/I fixed from Bars (values).", vbInformation
End Sub

