Attribute VB_Name = "modRepair"
Option Explicit

Public Sub RepairAndRebuild()
    On Error Resume Next
    Application.ShowFormulas = False
    Call RebuildBarsAll
End Sub

Public Sub HardTest_OneBlock()
    ' ??????Bars!B2:K2 ???? 7203 / Settings!B4 ???
    Dim wsB As Worksheet: Set wsB = ThisWorkbook.Sheets("Bars")
    wsB.Range("A2").Formula2 = "=RssChart(B2:K2,""7203"",TEXT(Settings!B4,""@""),20)"
End Sub
