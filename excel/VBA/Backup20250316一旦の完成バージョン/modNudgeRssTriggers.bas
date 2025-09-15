Attribute VB_Name = "modNudgeRssTriggers"
Option Explicit

'==== �⏕�F�擪�� =" �̌���͂������C�� ====
Private Sub StripAtIfNeeded(ByVal tgt As Range)
    On Error Resume Next
    Dim f As String: f = tgt.Formula2
    If Len(f) >= 2 Then If Left$(f, 2) = "='" Then tgt.Formula2 = "=" & Mid$(f, 3)
End Sub

'==== RssChart �p�̎��𐶐� ====
Private Function BuildRssChartFormula(ByVal headAddr As String, _
    ByVal codeCellAddr As String, ByVal foot As String) As String
    Dim fx As String
    fx = "=RssChart(" & headAddr & ",TEXT(" & codeCellAddr & ",""0""),""" & foot & """,20)"
    BuildRssChartFormula = fx
End Function

'==== �v�Z�ݒ�𐳏퉻�i���������C���j ====
Public Sub FixCalcAndRefresh()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then _
        Application.ActiveWindow.DisplayFormulas = False   '���������|�C���g
    Err.Clear
    On Error GoTo 0
    Application.CalculateFullRebuild
    Application.ScreenUpdating = True
End Sub

'==== �e�u���b�N�́u�g���K�[�v�Z���� RssChart ������������ ====
Public Sub NudgeRssTriggers()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsB As Worksheet: Set wsB = wb.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = wb.Sheets("Dashboard")
    Dim wsS As Worksheet: Set wsS = wb.Sheets("Settings")

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim blocks As Long: blocks = Application.Min(20, lastRow - 1)
    Dim r As Long, sc As Long
    Dim head As Range, codeCell As Range, trg As Range
    Dim foot As String: foot = CStr(wsS.Range("B4").Value)  '���� (��: "1M")

    For r = 1 To blocks
        sc = 2 + (r - 1) * 12                      'B=2, ���u���b�N�� +12 ��
        Set head = wsB.Range(wsB.Cells(2, sc), wsB.Cells(2, sc + 9)) 'B2:K2
        Set codeCell = wsD.Cells(r + 1, "A")       'Dashboard A��̖����R�[�h
        Set trg = wsB.Cells(2, sc - 1)             '�g���K�[��iA, M, �c�j
        trg.Formula2 = BuildRssChartFormula(head.Address(False, False), _
                                            codeCell.Address(False, False), foot)
        StripAtIfNeeded trg
    Next r
    Exit Sub
EH:
    MsgBox "NudgeRssTriggers error: " & Err.Description
End Sub

