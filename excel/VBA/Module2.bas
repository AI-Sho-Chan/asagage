Attribute VB_Name = "Module2"
Option Explicit

' 文字列/数値を日付に安全変換
Private Function DateFromCell(ByVal v As Variant) As Date
    On Error GoTo F
    If IsDate(v) Then DateFromCell = DateValue(v): Exit Function
    Dim s As String: s = CStr(v)
    s = Replace(Replace(Replace(s, "年", "/"), "月", "/"), "日", "")
    s = Replace(Replace(s, ".", "/"), "-", "/")
    If IsDate(s) Then DateFromCell = DateValue(s) Else DateFromCell = 0
    Exit Function
F:  DateFromCell = 0
End Function

' 数値化（カンマ等を除去）
Private Function ToDbl(ByVal v As Variant) As Double
    Dim s As String: s = Replace(CStr(v), ",", "")
    If IsNumeric(s) Then ToDbl = CDbl(s) Else ToDbl = 0#
End Function

' ブロック内の最新日付（今日以下）を返す
Private Function LatestDateInBlock(block As Range) As Date
    Dim r As Long, d As Date, m As Date: m = 0
    For r = 3 To block.Rows.Count
        d = DateFromCell(block.Cells(r, 4).Value) ' 4列目=日付
        If d <= Date And d > m Then m = d
    Next
    LatestDateInBlock = m
End Function

' ===== AVWAP（当日/指定日、無指定は最新日） =====
Public Function AVWAP_FromBlock(block As Range, Optional sess As Variant) As Variant
    On Error GoTo EH
    Dim target As Date, r As Long, d As Date, num As Double, den As Double
    If IsMissing(sess) Or Len(CStr(sess)) = 0 Then
        target = LatestDateInBlock(block)
    Else
        On Error Resume Next: target = DateValue(sess): On Error GoTo 0
        If target = 0 Then target = LatestDateInBlock(block)
    End If
    If target = 0 Then AVWAP_FromBlock = CVErr(xlErrNA): Exit Function

    For r = 3 To block.Rows.Count
        d = DateFromCell(block.Cells(r, 4).Value)
        If d = target Then
            num = num + ToDbl(block.Cells(r, 9).Value) * ToDbl(block.Cells(r, 10).Value) ' 終値×出来高
            den = den + ToDbl(block.Cells(r, 10).Value)
        End If
    Next
    AVWAP_FromBlock = IIf(den > 0, num / den, CVErr(xlErrNA))
    Exit Function
EH: AVWAP_FromBlock = CVErr(xlErrValue)
End Function

' ===== ATR(5)（当日/指定日、無指定は最新日） =====
Public Function ATR5_FromBlock(block As Range, Optional sess As Variant) As Variant
    On Error GoTo EH
    Dim target As Date: target = 0
    If IsMissing(sess) Or Len(CStr(sess)) = 0 Then
        target = LatestDateInBlock(block)
    Else
        On Error Resume Next: target = DateValue(sess): On Error GoTo 0
        If target = 0 Then target = LatestDateInBlock(block)
    End If
    If target = 0 Then ATR5_FromBlock = CVErr(xlErrNA): Exit Function

    Dim r As Long, d As Date, cnt As Long
    Dim pc As Variant, h As Double, l As Double, c As Double
    Dim tr As Double, sumtr As Double
    pc = Empty

    For r = block.Rows.Count To 3 Step -1
        d = DateFromCell(block.Cells(r, 4).Value)
        If d <> target Then If cnt = 0 Then GoTo NextRow Else Exit For
        h = ToDbl(block.Cells(r, 7).Value)
        l = ToDbl(block.Cells(r, 8).Value)
        c = ToDbl(block.Cells(r, 9).Value)
        If IsEmpty(pc) Then tr = h - l Else tr = WorksheetFunction.Max(h - l, Abs(h - pc), Abs(l - pc))
        sumtr = sumtr + tr: cnt = cnt + 1: pc = c
        If cnt = 5 Then Exit For
NextRow:
    Next
    ATR5_FromBlock = IIf(cnt > 0, sumtr / cnt, CVErr(xlErrNA))
    Exit Function
EH: ATR5_FromBlock = CVErr(xlErrValue)
End Function

