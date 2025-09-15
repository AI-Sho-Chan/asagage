Attribute VB_Name = "modAlerts"
Option Explicit

Public Sub CheckSignals()
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Dashboard")
    Dim last As Long, r As Long, dev As Double
    last = sh.Cells(sh.rows.Count, "A").End(xlUp).Row
    sh.Range("A2:L" & last).Interior.ColorIndex = xlNone

    For r = 2 To last
        dev = Nz(sh.Cells(r, "J").Value)
        If dev >= 0.6 Then
            sh.rows(r).Interior.Color = RGB(255, 235, 235) '”„‚èŒó•â
            Beep
        ElseIf dev <= -0.6 Then
            sh.rows(r).Interior.Color = RGB(235, 255, 235) '”ƒ‚¢Œó•â
            Beep
        End If
    Next
End Sub

