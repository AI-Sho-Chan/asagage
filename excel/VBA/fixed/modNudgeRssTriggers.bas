Attribute VB_Name = "modNudgeRssTriggers"
Option Explicit

Private Const TOP_ROW As Long = 2
Private Const COLS_PER_BLOCK As Long = 12
Private Const MAX_BLOCKS As Long = 20

Private Sub StripAtIfNeeded(ByVal tgt As Range)
    On Error Resume Next
    Dim f As String: f = tgt.Formula2
    If Len(f) >= 2 Then
        If Left$(f,2) = "=@"
        Then tgt.Formula2 = "=" & Mid$(f,3)
        End If
    End If
End Sub

Private Function Fx(ByVal headAddr As String, ByVal codeAddr As String, ByVal foot As String) As String
    Fx = "=RssChart(" & headAddr & ",IFERROR(TEXT(" & codeAddr & ",""0"")," & codeAddr & "&"""")" & ",""" & foot & """,20)"
End Function

Public Sub FixCalcAndRefresh()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.ShowFormulas = False
    Application.CalculateFullRebuild
    Application.ScreenUpdating = True
End Sub

Public Sub NudgeRssTriggers()
    On Error GoTo EH
    Dim wsB As Worksheet: Set wsB = ThisWorkbook.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("Dashboard")
    Dim foot As String: foot = ThisWorkbook.Sheets("Settings").Range("B4").Value

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    Dim blocks As Long: blocks = WorksheetFunction.Min(MAX_BLOCKS, WorksheetFunction.Max(0, lastRow - 1))

    Dim r As Long, sc As Long, hdr As Range, codeCell As Range, f As String
    For r = 2 To blocks + 1
        sc = 2 + (r - 2) * COLS_PER_BLOCK
        Set hdr     = wsB.Range(wsB.Cells(TOP_ROW, sc), wsB.Cells(TOP_ROW, sc + 9))
        Set codeCell= wsD.Cells(r, "A")
        f = Fx(hdr.Address(False, False), codeCell.Address(False, False), foot)
        With wsB.Cells(TOP_ROW, sc - 1)
            .NumberFormat = "General"
            .ClearContents
            .Formula2 = f
            StripAtIfNeeded wsB.Cells(TOP_ROW, sc - 1)
        End With
    Next r
    Exit Sub
EH:
    MsgBox "NudgeRssTriggers: " & Err.Description, vbExclamation
End Sub
