Attribute VB_Name = "modBarsSetup"
Option Explicit

Private Function HeaderArray() As Variant
    HeaderArray = Array("???","???","??","??","??","??","??","??","??","???")
End Function

Public Sub RebuildBarsAll()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsB As Worksheet: Set wsB = wb.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = wb.Sheets("Dashboard")

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count,"A").End(xlUp).Row
    Dim blocks As Long: blocks = Application.Min(20, Application.Max(0, lastRow-1))

    Application.ScreenUpdating = False
    If blocks > 0 Then wsB.Range(wsB.Cells(2,2), wsB.Cells(22, 12*blocks+1)).Clear

    Dim h As Variant: h = HeaderArray()
    Dim r As Long, sc As Long, k As Long
    For r = 2 To blocks + 1
        sc = 2 + (r-2)*12
        For k = 0 To UBound(h): wsB.Cells(2, sc+k).Value = h(k): Next k
    Next r

    Application.ScreenUpdating = True
    NudgeRssTriggers
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "RebuildBarsAll error: " & Err.Description, vbExclamation
End Sub
