
Attribute VB_Name = "QAE_Patch"
Option Explicit
Private Function EV(ByVal s As String) As Variant
    On Error GoTo e: EV = Application.Evaluate(s): Exit Function
e: EV = CVErr(xlErrNA)
End Function
Private Function NzNum(ByVal v As Variant) As Double
    If IsError(v) Or Not IsNumeric(v) Then NzNum = 0# Else NzNum = CDbl(v)
End Function
Private Function ArrMin(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long) As Double
    Dim i As Long, m As Double: m = 1E+308
    For i = lo To hi: If a(i) < m Then m = a(i)
    Next: If m = 1E+308 Then m = 0#: ArrMin = m
End Function
Private Function ArrMax(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long) As Double
    Dim i As Long, m As Double: m = -1E+308
    For i = lo To hi: If a(i) > m Then m = a(i)
    Next: If m = -1E+308 Then m = 0#: ArrMax = m
End Function
Private Function SafeNorm(ByVal x As Double, ByVal mn As Double, ByVal mx As Double) As Double
    If mx = mn Then SafeNorm = 0# Else SafeNorm = (x - mn) / (mx - mn)
End Function
Public Sub Patch_QAE()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long: last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    Dim Uv() As Double, Vv() As Double, Wv() As Double
    ReDim Uv(2 To last): ReDim Vv(2 To last): ReDim Wv(2 To last)
    Dim r As Long, code As String, px As Double
    Dim v20 As Double, bid As Double, ask As Double, nowp As Double, atr5 As Double
    For r = 2 To last
        code = Format$(D.Cells(r, "A").Value, "0")
        px   = Val(D.Cells(r, "C").Value)
        v20  = NzNum(EV("RssMarket(""" & code & """,""20日平均売買代金"")"))
        bid  = NzNum(EV("RssMarket(""" & code & """,""最良買気配"")"))
        ask  = NzNum(EV("RssMarket(""" & code & """,""最良売気配"")"))
        nowp = NzNum(EV("RssMarket(""" & code & """,""現在値"")"))
        atr5 = NzNum(EV("RssMarket(""" & code & """,""ATR(5)"")"))
        Uv(r) = v20
        If nowp > 0 And ask > 0 And bid > 0 Then Vv(r) = (ask - bid) / nowp Else Vv(r) = 1#
        Wv(r) = atr5 * IIf(nowp > 0, nowp, 0#)
        D.Cells(r, "U").Value = Uv(r)
        D.Cells(r, "V").Value = Vv(r)
        D.Cells(r, "W").Value = Wv(r)
        D.Cells(r, "X").Value = px
        D.Cells(r, "Y").Value = CStr(EV("RssMarket(""" & code & """,""市場区分"")"))
        D.Cells(r, "Z").Value = IIf(InStr(1, D.Cells(r, "Y").Value, "ETF", 1) Or InStr(1, D.Cells(r, "Y").Value, "REIT", 1), 1, 0)
        D.Cells(r, "AE").Value = (Uv(r) >= 2000000000#) And (px >= 500#) And (px <= 15000#) _
                                 And (Vv(r) <= 0.0025#) And (atr5 >= 1#)
    Next
    Dim Umin As Double, Umax As Double, Vmin As Double, Vmax As Double, Wmin As Double, Wmax As Double
    Umin = ArrMin(Uv, 2, last): Umax = ArrMax(Uv, 2, last)
    Vmin = ArrMin(Vv, 2, last): Vmax = ArrMax(Vv, 2, last)
    Wmin = ArrMin(Wv, 2, last): Wmax = ArrMax(Wv, 2, last)
    For r = 2 To last
        D.Cells(r, "AA").Value = SafeNorm(Uv(r), Umin, Umax)
        D.Cells(r, "AB").Value = SafeNorm(Wv(r), Wmin, Wmax)
        D.Cells(r, "AC").Value = SafeNorm(Vv(r), Vmin, Vmax)
        D.Cells(r, "AD").Value = IIf(D.Cells(r, "Z").Value = 1, -1000000000#, _
                              0.6# * D.Cells(r, "AA").Value + 0.5# * D.Cells(r, "AB").Value - 0.7# * D.Cells(r, "AC").Value)
    Next
End Sub
