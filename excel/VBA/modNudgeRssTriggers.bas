Attribute VB_Name = "modNudgeRssTriggers"
Option Explicit

Private Sub StripAtIfNeeded(ByVal tgt As Range)
    On Error Resume Next
    Dim f As String: f = tgt.Formula2
    If Len(f) >= 2 Then If Left$(f,2) = "=@" Then tgt.Formula2 = "=" & Mid$(f,3)
End Sub

Private Function GetSessionDateText() As String
    On Error Resume Next
    Dim v As Variant: v = ThisWorkbook.Sheets("Settings").Range("B5").Value
    If IsDate(v) Then GetSessionDateText = Format$(CDate(v),"yyyy/mm/dd") Else GetSessionDateText = Trim$(CStr(v))
End Function

Private Function BuildRssChartFormula(ByVal headAddr As String, ByVal codeCellAddr As String, ByVal foot As String, ByVal sessText As String) As String
    BuildRssChartFormula = "=RssChart(" & headAddr & ",TEXT(" & codeCellAddr & ",""0""),""" & foot & """,20)"
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
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsB As Worksheet: Set wsB = wb.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = wb.Sheets("Dashboard")
    Dim foot As String: foot = wb.Sheets("Settings").Range("B4").Value  ' ?: "1M"
    Dim sess As String: sess = GetSessionDateText()

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count,"A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim blocks As Long: blocks = Application.Min(20, lastRow-1)

    Application.ScreenUpdating = False
    Dim r As Long, sc As Long
    Dim headRange As Range, codeAddr As String, f As String

    For r = 2 To blocks + 1
        sc = 2 + (r-2)*12
        Set headRange = wsB.Range(wsB.Cells(2,sc), wsB.Cells(2,sc+9))  ' B:K ??
        codeAddr = wsD.Cells(r,"A").Address(False,False)               ' Dashboard!A*
        f = BuildRssChartFormula(headRange.Address(False,False), codeAddr, foot, sess)
        With wsB.Cells(2, sc - 1)  ' ???????????????????????
            .NumberFormat = "General"
            .ClearContents
            .Formula2 = f
            StripAtIfNeeded wsB.Cells(2, sc - 1)
        End With
    Next r

    FixCalcAndRefresh
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "NudgeRssTriggers error: " & Err.Description, vbExclamation
End Sub
