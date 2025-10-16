Attribute VB_Name = "Fix_O_to_AD"
Option Explicit

Public Sub Fix_O_to_AD()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)

    ' ���o������
    D.Range("O1").Value = "���m��(�~)"
    D.Range("P1").Value = "���P��"
    D.Range("Q1").Value = "�\�z�X���b�y�[�W:�G���g���["
    D.Range("R1").Value = "�\�z�X���b�y�[�W:����"
    D.Range("S1").Value = "�ŏI����"
    D.Range("T1").Value = "�������"            '��20�����ς͖��� �� �����l���̗p
    D.Range("U1").Value = "�X�v���b�h��"
    D.Range("V1").Value = "TR�~���i"
    D.Range("W1").Value = "���i"
    D.Range("X1").Value = "�s��敪"
    D.Range("Y1").Value = "���O�t���O"
    D.Range("Z1").Value = "z_������"
    D.Range("AA1").Value = "z_�{��"
    D.Range("AB1").Value = "z_�X�v���b�h"
    D.Range("AC1").Value = "����S"
    D.Range("AD1").Value = "����OK"

    ' @RssMarket �� RssMarket
    D.UsedRange.Replace What:="@RssMarket", Replacement:="RssMarket", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    Application.ScreenUpdating = False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' O: ���m��
            D.Cells(r, "O").Formula2 = "=IFERROR($C" & r & "-$L" & r & ",NA())"
            ' P: ���P�ʁi�P����񂪖��񋟂̂��ߊ���=100�j
            D.Cells(r, "P").Formula2 = "=IF($M" & r & "="""","""",100)"
            ' Q/R: �X���b�y�[�W
            D.Cells(r, "Q").Formula2 = _
              "=IF($M" & r & "="""","""",W_ENTRY_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),$C" & r & ",Settings!$B$22,Settings!$B$23,Settings!$B$31,Settings!$B$27))"
            D.Cells(r, "R").Formula2 = _
              "=IF($M" & r & "="""","""",W_EXIT_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""SELL"",""BUY""),$C" & r & ",W_QTY_BY_BUDGET_CLIPPED($C" & r & ",Settings!$B$22,Settings!$B$23,$A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),Settings!$B$31),Settings!$B$31,Settings!$B$27))"
            ' S: �ŏI����
            D.Cells(r, "S").Formula2 = "=IF(AND($R" & r & ">=Settings!$B$24,$AD" & r & "=TRUE),IF($J" & r & "<0,""GO SHORT"",""GO LONG""),""SKIP"")"
            ' T: �������
            D.Cells(r, "T").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�������""),0)"
            ' U: �X�v���b�h��
            D.Cells(r, "U").Formula2 = _
              "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z�l"")-RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z�l""))/RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),0)"
            ' V: TR�~���i
            D.Cells(r, "V").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"
            ' W: ���i
            D.Cells(r, "W").Formula2 = "=$C" & r
            ' X: �s��敪
            D.Cells(r, "X").Formula2 = _
              "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s�ꕔ����""),IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s�ꖼ��""),""""))"
            ' Y: ���O�t���O�iX�Q�Ƃŏz����j
            D.Cells(r, "Y").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",X" & r & ")),ISNUMBER(SEARCH(""REIT"",X" & r & "))),1,0)"

        End If
    Next

    ' Z,AA,AB �� z�X�R�A & AC/AD
    ' �͈͂�2:31�z��B�[��������B
    D.Range("Z2:Z31").Formula2 = _
      "=IFERROR((T2-AVERAGE($T$2:$T$31))/STDEV.P($T$2:$T$31),0)"
    D.Range("AA2:AA31").Formula2 = _
      "=IFERROR((V2-AVERAGE($V$2:$V$31))/STDEV.P($V$2:$V$31),0)"
    D.Range("AB2:AB31").Formula2 = _
      "=IFERROR((U2-AVERAGE($U$2:$U$31))/STDEV.P($U$2:$U$31),0)"
    D.Range("AC2:AC31").Formula2 = "=0.6*Z2+0.5*AA2-0.7*AB2"
    D.Range("AD2:AD31").Formula2 = _
      "=AND($W2>=500,$W2<=15000,$U2<=0.0025,$I2>=1,$Y2=0)"

    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "O?AD �̌��o���Ǝ����C�����܂����B", vbInformation
End Sub


