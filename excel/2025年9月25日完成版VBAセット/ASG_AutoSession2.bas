Attribute VB_Name = "ASG_AutoSession2"
Option Explicit
Public g_abort As Boolean

Public Sub Stop_AutoSession()
    g_abort = True
End Sub

Public Sub Run_AutoSession()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim s As Worksheet: Set s = Worksheets("Settings")
    Dim OUT As Worksheet
    On Error Resume Next: Set OUT = Worksheets("Signals"): On Error GoTo 0
    If OUT Is Nothing Then Set OUT = Worksheets.Add(After:=D): OUT.Name = "Signals"

    Dim startT As Date, minutes As Double
    startT = TimeValue(s.Range("B29").Value)     '�J�n ��: 09:00:00
    minutes = Val(s.Range("B37").Value)          '�Z�b�V�������� ��: 3 or 15
    If minutes <= 0 Then minutes = 3

    g_abort = False

    '�J�n�����܂őҋ@�i����O�͂����Ŏ~�܂��Č�����j
    Do While TimeValue(Now) < startT
        If g_abort Then Exit Sub
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop

    Dim tEnd As Date: tEnd = DateAdd("n", minutes, Now)

    Do While Now < tEnd
        If g_abort Then Exit Do
        Application.CalculateFull
        Call Recalc_Signal_And_Final_Once
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 5) '5�b�X�V
    Loop

    'Signals �o�́i���Ғl=O*P �~���j
    Call Export_Signals
    MsgBox "AutoSession�I��: " & Format(minutes, "0") & "��", vbInformation
End Sub

Private Sub Recalc_Signal_And_Final_Once()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim s As Worksheet: Set s = Worksheets("Settings")
    Dim last&, r&, j_th#, minS#
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)
    j_th = s.Range("B28").Value
    minS = s.Range("B24").Value

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' �\�蔭�����ƃX���b�v�X�V�iASG_Dashboard_Fix�̎����Ɠ������W�b�N��OK�j
            D.Cells(r, "P").Value = PlannedQty(NzNum(D.Cells(r, "C").Value))
            D.Cells(r, "Q").Value = SlipEst(CStr(D.Cells(r, "A").Value), IIf(D.Cells(r, "J").Value < 0, "BUY", "SELL"), NzNum(D.Cells(r, "C").Value), NzNum(D.Cells(r, "P").Value))
            D.Cells(r, "R").Value = SlipEst(CStr(D.Cells(r, "A").Value), IIf(D.Cells(r, "J").Value < 0, "SELL", "BUY"), NzNum(D.Cells(r, "C").Value), NzNum(D.Cells(r, "P").Value))

            D.Cells(r, "O").Formula2 = "=IFERROR($K" & r & "-($Q" & r & "+$R" & r & "),NA())"

            If IsNumeric(D.Cells(r, "J").Value) And Abs(D.Cells(r, "J").Value) >= j_th Then
                D.Cells(r, "M").Value = IIf(D.Cells(r, "J").Value < 0, "�G���g���[�V�O�i��", "�V���[�g�V�O�i��")
            Else
                D.Cells(r, "M").ClearContents
            End If

            D.Cells(r, "S").Formula2 = "=IF(AND($O" & r & ">=" & minS & ",$M" & r & "<>"""",$AD" & r & "),IF($J" & r & "<0,""GO LONG"",""GO SHORT""),""SKIP"")"
        End If
    Next
End Sub

Private Sub Export_Signals()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim OUT As Worksheet: Set OUT = Worksheets("Signals")
    Dim last&, r&, rowOut&
    OUT.Cells.Clear
    OUT.Range("A1:G1").Value = Array("�R�[�h", "������", "����", "�l�b�g���mO", "�\�蔭��P", "���Ғl(O*P)", "����")
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)
    rowOut = 2
    For r = 2 To last
        If D.Cells(r, "S").Value Like "GO *" Then
            OUT.Cells(rowOut, 1).Value = D.Cells(r, "A").Value
            OUT.Cells(rowOut, 2).Value = D.Cells(r, "B").Value
            OUT.Cells(rowOut, 3).Value = D.Cells(r, "S").Value
            OUT.Cells(rowOut, 4).Value = D.Cells(r, "O").Value
            OUT.Cells(rowOut, 5).Value = D.Cells(r, "P").Value
            OUT.Cells(rowOut, 6).FormulaR1C1 = "=RC[-2]*RC[-1]"
            OUT.Cells(rowOut, 7).Value = Now
            rowOut = rowOut + 1
        End If
    Next
    If rowOut > 2 Then OUT.Range("A2:G" & rowOut - 1).Sort Key1:=OUT.Range("F2"), Order1:=xlDescending, Header:=xlNo
End Sub

'? �ˑ��i�ȈՔŁj
Private Function PlannedQty(ByVal px As Double) As Double
    Dim budget#, lot#    'Settings B35=�\�Z, B36=�P��
    budget = Worksheets("Settings").Range("B35").Value
    lot = Worksheets("Settings").Range("B36").Value
    If px <= 0 Or lot <= 0 Then PlannedQty = 0: Exit Function
    PlannedQty = Int((budget / (px * lot))) * lot
End Function

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
        tick = 1#
        SlipEst = spread / 2 + levels * tick
    Else
        Dim beta#: beta = 0.2
        If adv <= 0 Then adv = 1
        SlipEst = spread / 2 + beta * (qty * px) / adv
    End If
End Function

Private Function EV(ByVal s As String) As Variant
    On Error GoTo e: EV = Application.Evaluate(s): Exit Function
e: EV = CVErr(xlErrNA)
End Function
Private Function NzNum(ByVal v As Variant) As Double
    If IsError(v) Or Not IsNumeric(v) Then NzNum = 0# Else NzNum = CDbl(v)
End Function

