Attribute VB_Name = "ASG_Universe_Min"
Option Explicit
Public Sub Build_From_UNIVERSE_EXTRA()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim U As Worksheet: Set U = Worksheets("UNIVERSE_EXTRA")
    Dim last As Long: last = U.Cells(U.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, n As Long, code As String
    D.Range("A2:A2000").ClearContents
    n = 0
    For r = 2 To last
        code = Format$(U.Cells(r, "A").Value, "0000")
        If Len(code) = 4 Then n = n + 1: D.Cells(n + 1, "A").Value = code
    Next
    ASG_Bars_Rebuild_FromDashboard
    FixHI_Run
    Patch_QAE
    MsgBox "Dashboard populated: " & n & " codes.", vbInformation
End Sub

