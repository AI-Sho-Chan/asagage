Attribute VB_Name = "ASG_Correctors"
Option Explicit

Public Sub Fix_U_V_W_X_Y()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)

    ' @ ��S����
    D.UsedRange.Replace What:="@RssMarket", Replacement:="RssMarket", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' U: ��������i�T���ς͖��� �� �����l���̗p�j
            D.Cells(r, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�������""),0)"
            ' V: �X�v���b�h����(�ŗǔ��C�z�l?�ŗǔ��C�z�l)/���ݒl
            D.Cells(r, "V").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z�l"")-RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z�l""))/RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),0)"
            ' W: TR�~���i�iI��TR=�~�AX�͉��i�j
            D.Cells(r, "W").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"
            D.Cells(r, "X").Formula2 = "=$C" & r
            ' Y: �s�ꕔ���́i������Ύs�ꖼ�̂ɕς���j
            D.Cells(r, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s�ꕔ����""),IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s�ꖼ��""),""""))"
            ' Z: ETF/REIT ���O
            D.Cells(r, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r & ")),ISNUMBER(SEARCH(""REIT"",Y" & r & "))),1,0)"
        End If
    Next
    Application.CalculateFull
    MsgBox "U/V/W/X/Y �𐳂����֐��ōĕ~�݁B", vbInformation
End Sub

Public Sub Fix_P_Q_R_and_Settings()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim S As Worksheet: Set S = Worksheets("Settings")
    Dim last As Long, r As Long
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)

    ' �W���������iI��TR=�~�Ȃ̂� 150��1.0�j
    S.Range("B22").Value = 1     '���m�W��
    S.Range("B23").Value = 1     '���،W��

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' P: �\�z�X���b�y�[�W(ENTRY)
            D.Cells(r, "P").Formula2 = _
              "=IF($M" & r & "="""","""",W_ENTRY_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),$C" & r & ",1,1,1,Settings!$B$27))"
            ' Q: �\�z�X���b�y�[�W(EXIT)
            D.Cells(r, "Q").Formula2 = _
              "=IF($M" & r & "="""","""",W_EXIT_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""SELL"",""BUY""),$C" & r & ",W_QTY_BY_BUDGET_CLIPPED($C" & r & ",1,1,$A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),1),1,Settings!$B$27))"
            ' R: �����m��
            D.Cells(r, "R").Formula2 = "=IF($M" & r & "="""","""",MAX(0,$O" & r & "-$P" & r & "-$Q" & r & "))"
        End If
    Next
    Application.CalculateFull
    MsgBox "P/Q/R ���C���BSettings!B22/B23 �� 1.0 �ɕύX�B", vbInformation
End Sub


