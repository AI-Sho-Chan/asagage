Attribute VB_Name = "modBarsSetup"
Option Explicit

Private Function HeaderArray() As Variant
    HeaderArray = Array("???","???/??","??","??","??","??","??","???")
End Function

Public Sub RebuildBarsAll()
    On Error GoTo EH
    Dim wsB As Worksheet, wsD As Worksheet
    Set wsB = ThisWorkbook.Sheets("Bars")
    Set wsD = ThisWorkbook.Sheets("Dashboard")

    Dim lastRow As Long, blocks As Long
    lastRow = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    blocks = Application.Max(0, lastRow - 1): If blocks > 20 Then blocks = 20

    Application.ScreenUpdating = False
    wsB.Range(wsB.Cells(2, 1), wsB.Cells(22, 12 * 20 + 1)).Clear

    Dim r As Long, sc As Long
    For r = 2 To blocks + 1
        sc = 2 + (r - 2) * 12
        wsB.Cells(2, sc).Resize(1, 8).Value = HeaderArray
    Next r

    Application.ScreenUpdating = True
    NudgeRssTriggers
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "RebuildBarsAll: " & Err.Description, vbExclamation
End Sub
