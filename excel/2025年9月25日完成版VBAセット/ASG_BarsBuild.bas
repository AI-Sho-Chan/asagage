Attribute VB_Name = "ASG_BarsBuild"
Option Explicit

'Bars: 見出し=1行目、データ式=2行目（ズレ無し）
Public Sub ASG_Bars_Rebuild_FromDashboard()
    Dim wsB As Worksheet, wsD As Worksheet
    Dim last As Long, r As Long, i As Long, sc As Long
    Dim code As String

    Set wsB = Worksheets("Bars")
    Set wsD = Worksheets("Dashboard")

    wsB.Cells.ClearContents

    ' 見出し（1行目）
    For i = 1 To 400
        sc = 1 + (i - 1) * 12   'A=1, M=13, Y=25, ...
        wsB.Cells(1, sc + 0).Value = "銘柄"
        wsB.Cells(1, sc + 1).Value = "名称"
        wsB.Cells(1, sc + 2).Value = "日付"
        wsB.Cells(1, sc + 3).Value = "時刻"
        wsB.Cells(1, sc + 4).Value = "始値"
        wsB.Cells(1, sc + 5).Value = "高値"
        wsB.Cells(1, sc + 6).Value = "安値"
        wsB.Cells(1, sc + 7).Value = "終値"
        wsB.Cells(1, sc + 8).Value = "出来高"
        wsB.Cells(1, sc + 9).Value = "売買代金"
    Next

    ' データ式（2行目に敷設）
    last = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    i = 0
    For r = 2 To last
        code = Format$(wsD.Cells(r, "A").Value, "0000")
        If Len(code) = 4 Then
            i = i + 1
            sc = 1 + (i - 1) * 12
            wsB.Cells(2, sc + 0).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""銘柄"")"
            wsB.Cells(2, sc + 1).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""名称"")"
            wsB.Cells(2, sc + 2).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""日付"")"
            wsB.Cells(2, sc + 3).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""時刻"")"
            wsB.Cells(2, sc + 4).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""始値"")"
            wsB.Cells(2, sc + 5).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""高値"")"
            wsB.Cells(2, sc + 6).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""安値"")"
            wsB.Cells(2, sc + 7).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""終値"")"
            wsB.Cells(2, sc + 8).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""出来高"")"
            wsB.Cells(2, sc + 9).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""売買代金"")"
        End If
    Next

    Application.CalculateFull
    MsgBox "Bars rebuilt for " & i & " codes.", vbInformation
End Sub

