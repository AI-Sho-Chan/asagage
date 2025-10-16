Attribute VB_Name = "ASG_Signal"
Option Explicit
' ���[���FJ�̕����ŕ�������AR(�����m) > Settings!B24 ���� AE=TRUE ��GO
Public Sub Make_Signal()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = 31
    Application.ScreenUpdating = False
    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' �V�O�i���iM�j
            If IsNumeric(D.Cells(r, "J").Value) Then
                If D.Cells(r, "J").Value < 0 Then
                    D.Cells(r, "M").Value = "�V���[�g�V�O�i��"
                Else
                    D.Cells(r, "M").Value = "�G���g���[�V�O�i��"
                End If
            Else
                D.Cells(r, "M").ClearContents
            End If
            ' �^�C���X�^���v�iN�j
            D.Cells(r, "N").Value = Now
            ' �\�z�X���b�y�[�W�ށiP,Q,R�j�͎��܂��͕ʓrUDF�ŉ^�p�B�����ł͖��ύX
            ' �ŏI����iS�j
            D.Cells(r, "S").Formula2 = _
                "=IF(AND($R" & r & ">=Settings!$B$24,$AE" & r & "=TRUE),IF($J" & r & "<0,""GO SHORT"",""GO LONG""),""SKIP"")"
        End If
    Next
    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "M/N/S ���X�V���܂����B", vbInformation
End Sub


