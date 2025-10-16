Attribute VB_Name = "ASG_Repair_HI"
Option Explicit

Public Sub Repair_HI_And_Test()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long, code As String
    last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then last = 31

    Application.ScreenUpdating = False

    ' 1) H/I �̏����Ɠ��e���������i�����񁨐����ɖ߂��j
    With D.Range("H2:I" & last)
        .NumberFormat = "General"
        .Replace What:="'=", Replacement:="=", LookAt:=xlPart
        .Replace What:="@RssMarket", Replacement:="RssMarket", LookAt:=xlPart
        .ClearContents
    End With

    ' 2) H/I �̐������ĕ~�݁iRssMarket ���j
    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""����VWAP""),NA())"
            D.Cells(r, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""ATR(5)""),NA())"
        End If
    Next

    ' 3) J/K/L ���Čv�Z����ۏ�
    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"
        End If
    Next

    Application.CalculateFull
    Application.ScreenUpdating = True

    ' 4) RSS�A�h�C���a�ʃe�X�g�iA2�̃R�[�h��1�񂾂��j
    code = Trim$(CStr(D.Cells(2, "A").Value))
    If Len(code) = 0 Then
        MsgBox "A2 ����ł��B�܂��R�[�h�����Ă��������B", vbExclamation: Exit Sub
    End If
    Dim t1 As Variant, t2 As Variant
    t1 = Application.Evaluate("RssMarket(""" & code & """,""����VWAP"")")
    t2 = Application.Evaluate("RssMarket(""" & code & """,""ATR(5)"")")

    If IsError(t1) Or IsError(t2) Then
        MsgBox "RSS�������܂��̓t�B�[���h���s��v�B�u=RssMarket(""" & code & """,""����VWAP"")�v��C�ӃZ���ɒ��ł����Ēl���Ԃ邩�m�F�B�Ԃ�Ȃ����MarketSpeed II�̃A�h�C����L�������AExcel�ċN���B", vbExclamation
    Else
        MsgBox "H/I �����ĕ~�݂��܂����B�e�X�g�l: VWAP=" & t1 & " / ATR5=" & t2, vbInformation
    End If
End Sub

