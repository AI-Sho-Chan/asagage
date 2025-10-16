Attribute VB_Name = "ASG_AutoCalibrate"
Option Explicit

' ���̂������ۂɒl���Ԃ鍀�ږ��������^�C������
Private Function ResolveField(ByVal code As String, ByRef candidates() As String) As String
    Dim i As Long, v As Variant
    For i = LBound(candidates) To UBound(candidates)
        v = Application.Evaluate("RssMarket(""" & code & """,""" & candidates(i) & """)")
        If Not IsError(v) Then
            If IsNumeric(v) Or Len(CStr(v)) > 0 Then
                ResolveField = candidates(i)
                Exit Function
            End If
        End If
    Next
    ResolveField = ""
End Function

' H/I �̎��� �g���������ږ��h �ŕ~�݂������iA2:A31 ���Ώہj
Public Sub Calibrate_And_Fill_HI()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim code As String
    code = Trim$(CStr(D.Cells(2, "A").Value))
    If Len(code) = 0 Then MsgBox "A2 ����ł��B�R�[�h�����Ă��������B", vbExclamation: Exit Sub

    ' ��⃊�X�g�i���@�ŕԂ���̂�T���j
    Dim vwapCands() As String, atrCands() As String
    vwapCands = Split("VWAP,����VWAP,�o�������d����,�����o�������d����,�����o�������d���ω��i,���d����", ",")
    atrCands = Split("ATR(5),ATR 5,�A�x���[�W�E�g�D���[�E�����W(5),ATR,ATR(14),ATR % (14) 1��,ATR��(14) 1��", ",")

    Dim vwapKey As String, atrKey As String
    vwapKey = ResolveField(code, vwapCands)
    atrKey = ResolveField(code, atrCands)

    If vwapKey = "" Then
        MsgBox "VWAP�n�̍��ږ���������܂���B�蓮�Ő����������Ă��������B", vbCritical
        Exit Sub
    End If
    If atrKey = "" Then
        MsgBox "ATR�n�̍��ږ���������܂���B�蓮�Ő����������Ă��������B", vbCritical
        Exit Sub
    End If

    ' �ݒu
    Dim last As Long, r As Long
    last = 31  ' 30�����Œ�^�p
    Application.ScreenUpdating = False

    ' H/I ���N���A���Ċm���ɐ�����
    With D.Range("H2:I" & last)
        .NumberFormat = "General"
        .Replace What:="'=", Replacement:="=", LookAt:=xlPart
        .Replace What:="@RssMarket", Replacement:="RssMarket", LookAt:=xlPart
        .ClearContents
    End With

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""" & vwapKey & """),NA())"
            D.Cells(r, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""" & atrKey & """),NA())"
            ' �ˑ�����ĕ~��
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"
        End If
    Next

    Application.CalculateFull

    MsgBox "VWAP=""" & vwapKey & """ / ATR=""" & atrKey & """ �� H/I ���ĕ~�݂��܂����B", vbInformation
End Sub

