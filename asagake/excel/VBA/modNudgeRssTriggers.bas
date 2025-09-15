Attribute VB_Name = "modNudgeRssTriggers"
Option Explicit

Private Const TOP_ROW As Long = 2
Private Const COLS_PER_BLOCK As Long = 12
Private Const MAX_BLOCKS As Long = 20

Private Sub StripAtIfNeeded(ByVal tgt As Range)
    On Error Resume Next
    If Left$(tgt.Formula2, 2) = "=@" Then tgt.Formula2 = "=" & Mid$(tgt.Formula2, 3)
End Sub

Private Function BuildRssChartFormula( _
    ByVal headAddr As String, _
    ByVal codeCellAddr As String, _
    ByVal foot As String) As String

    Dim fx As String
    fx = "=RssChart(" & headAddr & ",TEXT(" & codeCellAddr & ",""0""),""" & foot & """,20)"
    BuildRssChartFormula = fx
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
    Dim wsB As Worksheet, wsD As Worksheet, wsS As Worksheet
    Set wsB = ThisWorkbook.Sheets("Bars")
    Set wsD = ThisWorkbook.Sheets("Dashboard")
    Set wsS = ThisWorkbook.Sheets("Settings")

    Dim foot As String: foot = Trim$(wsS.Range("B4").Value)
    If Len(foot) = 0 Then foot = "1M"

    Dim lastRow As Long, blocks As Long
    lastRow = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    blocks = Application.Min(MAX_BLOCKS, Application.Max(0, lastRow - (TOP_ROW - 1)))
    If blocks = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Dim r As Long, sc As Long, rngHead As Range, fx As String
    For r = TOP_ROW To TOP_ROW + blocks - 1
        sc = 2 + (r - TOP_ROW) * COLS_PER_BLOCK
        Set rngHead = wsB.Range(wsB.Cells(2, sc), wsB.Cells(2, sc + 9))
        fx = BuildRssChartFormula(rngHead.Address(False, False), wsD.Cells(r, "A").Address(False, False), foot)
        With wsB.Cells(2, sc)
            .NumberFormat = "General"
            .ClearContents
            .Formula2 = fx
            StripAtIfNeeded wsB.Cells(2, sc)
        End With
    Next r
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "NudgeRssTriggers: " & Err.Description, vbExclamation
End Sub
