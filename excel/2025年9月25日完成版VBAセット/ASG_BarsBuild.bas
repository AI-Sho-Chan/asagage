Attribute VB_Name = "ASG_BarsBuild"
Option Explicit

'Bars: ���o��=1�s�ځA�f�[�^��=2�s�ځi�Y�������j
Public Sub ASG_Bars_Rebuild_FromDashboard()
    Dim wsB As Worksheet, wsD As Worksheet
    Dim last As Long, r As Long, i As Long, sc As Long
    Dim code As String

    Set wsB = Worksheets("Bars")
    Set wsD = Worksheets("Dashboard")

    wsB.Cells.ClearContents

    ' ���o���i1�s�ځj
    For i = 1 To 400
        sc = 1 + (i - 1) * 12   'A=1, M=13, Y=25, ...
        wsB.Cells(1, sc + 0).Value = "����"
        wsB.Cells(1, sc + 1).Value = "����"
        wsB.Cells(1, sc + 2).Value = "���t"
        wsB.Cells(1, sc + 3).Value = "����"
        wsB.Cells(1, sc + 4).Value = "�n�l"
        wsB.Cells(1, sc + 5).Value = "���l"
        wsB.Cells(1, sc + 6).Value = "���l"
        wsB.Cells(1, sc + 7).Value = "�I�l"
        wsB.Cells(1, sc + 8).Value = "�o����"
        wsB.Cells(1, sc + 9).Value = "�������"
    Next

    ' �f�[�^���i2�s�ڂɕ~�݁j
    last = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    i = 0
    For r = 2 To last
        code = Format$(wsD.Cells(r, "A").Value, "0000")
        If Len(code) = 4 Then
            i = i + 1
            sc = 1 + (i - 1) * 12
            wsB.Cells(2, sc + 0).Formula2 = "=RssChart(""" & code & """,""������"",,,""����"")"
            wsB.Cells(2, sc + 1).Formula2 = "=RssChart(""" & code & """,""������"",,,""����"")"
            wsB.Cells(2, sc + 2).Formula2 = "=RssChart(""" & code & """,""������"",,,""���t"")"
            wsB.Cells(2, sc + 3).Formula2 = "=RssChart(""" & code & """,""������"",,,""����"")"
            wsB.Cells(2, sc + 4).Formula2 = "=RssChart(""" & code & """,""������"",,,""�n�l"")"
            wsB.Cells(2, sc + 5).Formula2 = "=RssChart(""" & code & """,""������"",,,""���l"")"
            wsB.Cells(2, sc + 6).Formula2 = "=RssChart(""" & code & """,""������"",,,""���l"")"
            wsB.Cells(2, sc + 7).Formula2 = "=RssChart(""" & code & """,""������"",,,""�I�l"")"
            wsB.Cells(2, sc + 8).Formula2 = "=RssChart(""" & code & """,""������"",,,""�o����"")"
            wsB.Cells(2, sc + 9).Formula2 = "=RssChart(""" & code & """,""������"",,,""�������"")"
        End If
    Next

    Application.CalculateFull
    MsgBox "Bars rebuilt for " & i & " codes.", vbInformation
End Sub

