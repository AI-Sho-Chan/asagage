Attribute VB_Name = "modBars"
Option Explicit

Private Const MAX_BLOCKS As Long = 20   '���� 20 �u���b�N�܂�

Private Function HeaderArray() As Variant
    HeaderArray = Array("������", "�s�ꖼ/����", "���t", "����", "�n�l", "���l", "���l", "�I�l", "�o����")
End Function

Public Sub RebuildBarsAll()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsB As Worksheet: Set wsB = wb.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = wb.Sheets("Dashboard")

    Dim sess As String: sess = Trim$(CStr(wsD.Range("B4").Value))
    If Len(sess) = 0 Then sess = "1M"

    Application.ScreenUpdating = False
    wsB.Range(wsB.Cells(2, 1), wsB.Cells(22, 12 * MAX_BLOCKS + 1)).Clear

    Dim headers As Variant: headers = HeaderArray()
    Dim r As Long, sc As Long, k As Long
    For r = 2 To 21 ' Dashboard!A2:A21 �� 20 �����Ƃ݂Ȃ�
        Dim code As String: code = CleanTicker(wsD.Cells(r, 1).Value)
        If Len(code) = 0 Then Exit For

        sc = 2 + (r - 2) * 12                  '�u���b�N�擪��iB=2�j
        With wsB
            ' �w�b�_�[
            For k = LBound(headers) To UBound(headers)
                .Cells(2, sc + k).Value = headers(k)
            Next
            ' ���[�Z���iA, M, Y, �c�j�� RssChart ������
            .Cells(2, sc - 1).Formula = "=RssChart(B2:K2,""" & code & """,""" & sess & """,20)"
        End With
    Next

    Application.CalculateFull
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "RebuildBarsAll failed: " & Err.Description, vbExclamation
End Sub

' #VALUE! ���c��ۂ̓��꒼���i�K�v���̂݁j
Public Sub NudgeRssTriggers()
    Dim wsB As Worksheet: Set wsB = ThisWorkbook.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("Dashboard")
    Dim sess As String: sess = Trim$(CStr(wsD.Range("B4").Value))
    If Len(sess) = 0 Then sess = "1M"

    Dim r As Long, sc As Long, code As String
    For r = 2 To 21
        code = CleanTicker(wsD.Cells(r, 1).Value)
        If Len(code) = 0 Then Exit For
        sc = 2 + (r - 2) * 12
        wsB.Cells(2, sc - 1).Formula = ""
        wsB.Cells(2, sc - 1).Formula = "=RssChart(B2:K2,""" & code & """,""" & sess & """,20)"
    Next
    Application.CalculateFull
End Sub

