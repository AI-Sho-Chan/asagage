
Attribute VB_Name = "ASG_BarsBuild"
Option Explicit
Public Sub ASG_Bars_Rebuild_FromDashboard()
    Dim wsB As Worksheet: Set wsB = Worksheets("Bars")
    Dim wsD As Worksheet: Set wsD = Worksheets("Dashboard")
    Dim last As Long: last = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, i As Long, sc As Long, code As String
    wsB.Cells.ClearContents
    For i = 1 To 400
        sc = 2 + (i - 1) * 12
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
    i = 0
    For r = 2 To last
        code = Format$(wsD.Cells(r, "A").Value, "0000")
        If Len(code) = 4 Then
            i = i + 1: sc = 2 + (i - 1) * 12
            wsB.Cells(2, sc - 1).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""銘柄"")"
            wsB.Cells(3, sc + 0).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""銘柄"")"
            wsB.Cells(3, sc + 1).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""名称"")"
            wsB.Cells(3, sc + 2).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""日付"")"
            wsB.Cells(3, sc + 3).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""時刻"")"
            wsB.Cells(3, sc + 4).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""始値"")"
            wsB.Cells(3, sc + 5).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""高値"")"
            wsB.Cells(3, sc + 6).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""安値"")"
            wsB.Cells(3, sc + 7).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""終値"")"
            wsB.Cells(3, sc + 8).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""出来高"")"
            wsB.Cells(3, sc + 9).Formula2 = "=RssChart(""" & code & """,""日中足"",,,""売買代金"")"
        End If
    Next
    Application.CalculateFull
    MsgBox "Bars rebuilt for " & i & " codes.", vbInformation
End Sub
