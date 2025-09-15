Attribute VB_Name = "modBarsSetup"
'--- modBarsSetup.bas ---
Option Explicit
Private Const FOOT As String = "1M"   ' OneBlock で通った表記に（"1M" / "1分" など）

Public Sub RebuildBarsAll()
    Dim wsB As Worksheet: Set wsB = Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = Sheets("Dashboard")

    Dim last As Long, n As Long
    last = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    n = Application.Min(20, WorksheetFunction.Max(1, last - 1))  ' A2..A(last)

    Dim headers: headers = Array("銘柄名称", "市場名称", "足種", "日付", "時刻", "始値", "高値", "安値", "終値", "出来高")

    Application.ScreenUpdating = False
    wsB.Range(wsB.Cells(2, 1), wsB.Cells(22, 12 * 20 + 1)).Clear

    Dim i As Long, sc As Long, k As Long, rngHead As Range
    Dim code As String, addr As String

    For i = 2 To n + 1
        sc = 2 + (i - 2) * 12
        Set rngHead = wsB.Range(wsB.Cells(2, sc), wsB.Cells(2, sc + 9))
        For k = 0 To 9: rngHead.Cells(1, k + 1).Value = headers(k): Next k

        ' A列の銘柄コードをVBAでクリーン化
        code = CleanTicker(wsD.Cells(i, "A").Value2)
        If Len(code) = 0 Then GoTo NextI  ' 空行はスキップ

        ' 第1引数は同一シート参照（シート名を付けない）
        addr = rngHead.Address(False, False)

        With wsB.Cells(2, sc - 1)
            .NumberFormat = "General"
            .ClearContents
            .Formula2 = "=RssChart(" & addr & ",""" & code & """,""" & FOOT & """,20)"
        End With
NextI:
    Next i

    Application.ScreenUpdating = True
    Application.CalculateFull
End Sub
