Attribute VB_Name = "FixHI_Min"
Option Explicit

'--- helpers ---
Private Function DayKey(ByVal v As Variant) As Long
    On Error GoTo fail
    If IsDate(v) Then DayKey = CLng(CDbl(CDate(v))): Exit Function
    If IsNumeric(v) Then DayKey = CLng(CDbl(v)): Exit Function
fail: DayKey = 0
End Function

Private Function BlockLastRow(ByVal sc As Long) As Long
    Dim ws As Worksheet: Set ws = Worksheets("Bars")
    Dim j As Long, lr As Long, mx As Long
    mx = 3
    For j = sc To sc + 9
        lr = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
        If lr > mx Then mx = lr
    Next
    BlockLastRow = mx
End Function

Private Function FindBlockCol(ByVal code As String) As Long
    Dim ws As Worksheet: Set ws = Worksheets("Bars")
    Dim i As Long, sc As Long, fx As String
    For i = 1 To 400
        sc = 2 + (i - 1) * 12
        fx = CStr(ws.Cells(2, sc - 1).Formula2)
        If Len(fx) = 0 Then Exit For
        fx = Replace(fx, "=@RssChart", "=RssChart")
        If InStr(1, fx, """" & code & """", vbTextCompare) > 0 Then
            FindBlockCol = sc
            Exit Function
        End If
    Next
End Function

Private Function VWAP_FromBars(ByVal sc As Long) As Variant
    Dim ws As Worksheet: Set ws = Worksheets("Bars")
    Const oD As Long = 3, oC As Long = 8, oV As Long = 9
    Dim cD As Long, cC As Long, cV As Long
    cD = sc + oD: cC = sc + oC: cV = sc + oV
    Dim last As Long, r As Long, vSum As Double, pvSum As Double
    Dim kt As Long, haveToday As Boolean, k As Long, kLast As Long
    last = BlockLastRow(sc): kt = CLng(Date)
    For r = 3 To last
        k = DayKey(ws.Cells(r, cD).Value2)
        If k = kt Then
            haveToday = True
            If IsNumeric(ws.Cells(r, cV).Value2) And IsNumeric(ws.Cells(r, cC).Value2) Then
                vSum = vSum + CDbl(ws.Cells(r, cV).Value2)
                pvSum = pvSum + CDbl(ws.Cells(r, cV).Value2) * CDbl(ws.Cells(r, cC).Value2)
            End If
        End If
    Next
    If haveToday And vSum > 0 Then VWAP_FromBars = pvSum / vSum: Exit Function
    vSum = 0: pvSum = 0: kLast = 0
    For r = last To 3 Step -1
        k = DayKey(ws.Cells(r, cD).Value2)
        If k > 0 Then kLast = k: Exit For
    Next
    If kLast = 0 Then VWAP_FromBars = CVErr(xlErrNA): Exit Function
    For r = 3 To last
        If DayKey(ws.Cells(r, cD).Value2) = kLast Then
            If IsNumeric(ws.Cells(r, cV).Value2) And IsNumeric(ws.Cells(r, cC).Value2) Then
                vSum = vSum + CDbl(ws.Cells(r, cV).Value2)
                pvSum = pvSum + CDbl(ws.Cells(r, cV).Value2) * CDbl(ws.Cells(r, cC).Value2)
            End If
        End If
    Next
    If vSum > 0 Then VWAP_FromBars = pvSum / vSum Else VWAP_FromBars = CVErr(xlErrNA)
End Function

Private Function ATR5_FromBars(ByVal sc As Long) As Variant
    Dim ws As Worksheet: Set ws = Worksheets("Bars")
    Const oD As Long = 3, oH As Long = 6, oL As Long = 7, oC As Long = 8
    Dim cD As Long, cH As Long, cL As Long, cC As Long
    cD = sc + oD: cH = sc + oH: cL = sc + oL: cC = sc + oC
    Dim last As Long, r As Long, cnt As Long, trSum As Double, prevC As Double
    Dim kt As Long, k As Long, kTarget As Long, haveToday As Boolean
    last = BlockLastRow(sc): kt = CLng(Date)
    kTarget = 0
    For r = last To 3 Step -1
        k = DayKey(ws.Cells(r, cD).Value2)
        If k = kt Then kTarget = kt: haveToday = True: Exit For
        If kTarget = 0 And k > 0 Then kTarget = k
    Next
    If kTarget = 0 Then ATR5_FromBars = CVErr(xlErrNA): Exit Function
    For r = last To 3 Step -1
        If DayKey(ws.Cells(r, cD).Value2) <> kTarget Then
            If haveToday Then Exit For Else GoTo SkipRow
        End If
        If IsNumeric(ws.Cells(r, cH).Value2) And IsNumeric(ws.Cells(r, cL).Value2) And IsNumeric(ws.Cells(r, cC).Value2) Then
            Dim hi As Double, lo As Double, cls As Double, tr As Double
            hi = CDbl(ws.Cells(r, cH).Value2)
            lo = CDbl(ws.Cells(r, cL).Value2)
            cls = CDbl(ws.Cells(r, cC).Value2)
            If cnt = 0 Then prevC = cls
            tr = Application.Max(hi - lo, Abs(hi - prevC), Abs(lo - prevC))
            trSum = trSum + tr: cnt = cnt + 1: prevC = cls
            If cnt = 5 Then Exit For
        End If
SkipRow:
    Next
    If cnt > 0 Then ATR5_FromBars = trSum / cnt Else ATR5_FromBars = CVErr(xlErrNA)
End Function

'--- main ---
Public Sub FixHI_Run()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long: last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, code As String, sc As Long
    Dim vH As Variant, vI As Variant, hRM As Variant, iRM As Variant

    Application.ScreenUpdating = False
    For r = 2 To last
        code = CStr(D.Cells(r, "A").Value)
        If Len(code) = 4 Then
            sc = FindBlockCol(code)
            If sc > 0 Then
                vH = VWAP_FromBars(sc)
                vI = ATR5_FromBars(sc)
            Else
                vH = CVErr(xlErrNA)
                vI = CVErr(xlErrNA)
            End If

            ' RssMarket フォールバック
            hRM = Application.Evaluate("RssMarket(""" & code & """,""当日VWAP"")")
            iRM = Application.Evaluate("RssMarket(""" & code & """,""ATR(5)"")")
            If IsError(vH) Or Not IsNumeric(vH) Then If Not IsError(hRM) Then vH = hRM
            If IsError(vI) Or Not IsNumeric(vI) Then If Not IsError(iRM) Then vI = iRM

            If Not IsError(vH) Then D.Cells(r, "H").Value = vH
            If Not IsError(vI) Then D.Cells(r, "I").Value = vI
        End If
    Next
    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "H/I fixed (Bars or RssMarket).", vbInformation
End Sub

