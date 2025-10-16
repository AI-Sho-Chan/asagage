Attribute VB_Name = "ASG_Dashboard_Fix"
Option Explicit

Private Function EV(ByVal S As String) As Variant
    On Error GoTo e: EV = Application.Evaluate(S): Exit Function
e: EV = CVErr(xlErrNA)
End Function

'? ����e�B�b�N�i�ȈՁj
Private Function TickSize(ByVal px As Double) As Double
    If px < 1000# Then TickSize = 1# Else TickSize = 1#
End Function

'? �\�蔭�����ʁi�\�Z/���i/�P�����j
Private Function PlannedQty(ByVal px As Double) As Double
    Dim budget#, lot#
    budget = Worksheets("Settings").Range("B35").Value
    lot = Worksheets("Settings").Range("B36").Value
    If px <= 0 Or lot <= 0 Then PlannedQty = 0: Exit Function
    PlannedQty = Int((budget / (px * lot))) * lot
End Function

'? �X���b�y�[�W���v�i�ꒆ=�A��O=�����ʉ�A�j
Private Function SlipEst(ByVal code As String, ByVal side As String, ByVal px As Double, ByVal qty As Double) As Double
    Dim ask#, bid#, aSz#, bSz#, spread#, adv#
    ask = NzNum(EV("RssMarket(""" & code & """,""�ŗǔ��C�z�l"")"))
    bid = NzNum(EV("RssMarket(""" & code & """,""�ŗǔ��C�z�l"")"))
    aSz = NzNum(EV("RssMarket(""" & code & """,""�ŗǔ��C�z����"")"))
    bSz = NzNum(EV("RssMarket(""" & code & """,""�ŗǔ��C�z����"")"))
    spread = Application.Max(0, ask - bid)
    adv = NzNum(EV("RssMarket(""" & code & """,""�o����"")")) * px
    If ask > 0 And bid > 0 And ((aSz > 0) Or (bSz > 0)) Then
        Dim need#, baseSz#, levels#, tick#
        If UCase$(side) = "BUY" Then baseSz = aSz Else baseSz = bSz
        need = Application.Max(0, qty - baseSz)
        levels = need / Application.Max(1, baseSz)
        tick = TickSize(px)
        SlipEst = spread / 2 + levels * tick
    Else
        Dim beta#: beta = 0.2
        If adv <= 0 Then adv = 1
        SlipEst = spread / 2 + beta * (qty * px) / adv
    End If
End Function

Private Function NzNum(ByVal v As Variant) As Double
    If IsError(v) Or Not IsNumeric(v) Then NzNum = 0# Else NzNum = CDbl(v)
End Function

Public Sub Apply_All_Fixes()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim S As Worksheet: Set S = Worksheets("Settings")
    Dim r&, last&, code$, k_tp#, k_sl#, j_th#, minS#

    k_tp = S.Range("B22").Value: k_sl = S.Range("B23").Value
    j_th = S.Range("B28").Value: minS = S.Range("B24").Value

    ' ���o�����m��i�ꗉ����j
    D.Range("O1").Value = "�l�b�g���m��(�~/��)" 'K ? (Q+R)
    D.Range("P1").Value = "�\�蔭������[��]"
    D.Range("Q1").Value = "�\�z�X���b�y�[�W:�G���g���[[�~/��]"
    D.Range("R1").Value = "�\�z�X���b�y�[�W:����[�~/��]"
    D.Range("S1").Value = "�ŏI����"
    D.Range("T1").Value = "�������"
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

    ' @ ������
    D.UsedRange.Replace "@RssMarket", "RssMarket", xlPart

    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)
    Application.ScreenUpdating = False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            code = Format$(D.Cells(r, "A").Value, "0")

            ' H: VWAP
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�o�������d����""),NA())"
            ' I: TR
            D.Cells(r, "I").Formula2 = "=IFERROR(MAX($E" & r & "-$F" & r & _
                ",ABS($E" & r & "-RssMarket(TEXT($A" & r & ",""0""),""�O���I�l""))," & _
                "ABS($F" & r & "-RssMarket(TEXT($A" & r & ",""0""),""�O���I�l""))),NA())"
            ' J: ����
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            ' K/L: �����I�ȕ��iTR�ƍ����ŏ������j
            D.Cells(r, "K").Formula2 = "=IFERROR(MIN(Settings!$B$22*$I" & r & ",0.8*($E" & r & "-$F" & r & ")),NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR(MIN(Settings!$B$23*$I" & r & ",1.0*($E" & r & "-$F" & r & ")),NA())"

            ' W=���i
            D.Cells(r, "W").Formula2 = "=$C" & r
            ' X=�s��敪
            D.Cells(r, "X").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s�ꕔ����""),IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s�ꖼ��""),""""))"
            ' Y=���O
            D.Cells(r, "Y").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",X" & r & ")),ISNUMBER(SEARCH(""REIT"",X" & r & "))),1,0)"
            ' T,U,V
            D.Cells(r, "T").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�������""),0)"
            D.Cells(r, "U").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z�l"")-RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z�l""))/RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),0)"
            D.Cells(r, "V").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"

            ' P=�\�蔭������
            D.Cells(r, "P").Value = PlannedQty(NzNum(D.Cells(r, "C").Value))

            ' Q/R �X���b�y�[�W�i�l�Ŗ��߂�j
            D.Cells(r, "Q").Value = SlipEst(code, IIf(D.Cells(r, "J").Value < 0, "BUY", "SELL"), NzNum(D.Cells(r, "C").Value), NzNum(D.Cells(r, "P").Value))
            D.Cells(r, "R").Value = SlipEst(code, IIf(D.Cells(r, "J").Value < 0, "SELL", "BUY"), NzNum(D.Cells(r, "C").Value), NzNum(D.Cells(r, "P").Value))

            ' O=�l�b�g���m��
            D.Cells(r, "O").Formula2 = "=IFERROR($K" & r & "-($Q" & r & "+$R" & r & "),NA())"

            ' M=�V�O�i���iJ臒l�j
            If IsNumeric(D.Cells(r, "J").Value) Then
                If Abs(D.Cells(r, "J").Value) >= j_th Then
                    D.Cells(r, "M").Value = IIf(D.Cells(r, "J").Value < 0, "�G���g���[�V�O�i��", "�V���[�g�V�O�i��")
                Else
                    D.Cells(r, "M").ClearContents
                End If
            Else
                D.Cells(r, "M").ClearContents
            End If

            ' S=�ŏI����i�l�b�g���m�Ə����j
            D.Cells(r, "S").Formula2 = "=IF(AND($O" & r & ">=" & minS & ",$M" & r & "<>"""",$AD" & r & "),IF($J" & r & "<0,""GO LONG"",""GO SHORT""),""SKIP"")"
        End If
    Next

    ' Z..AD�iz�W�����Ə���OK�j
    D.Range("Z2:Z31").Formula2 = "=IFERROR((T2-AVERAGE($T$2:$T$31))/STDEV.P($T$2:$T$31),0)"
    D.Range("AA2:AA31").Formula2 = "=IFERROR((V2-AVERAGE($V$2:$V$31))/STDEV.P($V$2:$V$31),0)"
    D.Range("AB2:AB31").Formula2 = "=IFERROR((U2-AVERAGE($U$2:$U$31))/STDEV.P($U$2:$U$31),0)"
    D.Range("AC2:AC31").Formula2 = "=0.6*Z2+0.5*AA2-0.7*AB2"
    D.Range("AD2:AD31").Formula2 = "=AND($W2>=500,$W2<=15000,$U2<=0.0025,$I2>=" & S.Range("B26").Value & ",$Y2=0)"

    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "Dashboard�C�������F8306�܂ߔ񌻎����������I���ցAJ臒l�A���A�l�b�g���m(O)�A�ŏI����(S)�B", vbInformation
End Sub


