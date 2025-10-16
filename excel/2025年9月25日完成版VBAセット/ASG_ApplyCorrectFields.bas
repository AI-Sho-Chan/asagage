Attribute VB_Name = "ASG_ApplyCorrectFields"
Option Explicit

' Dashboard��H/I�ƈˑ�����u�������֐��v�ɏC��
' H = �o�������d���ρAI = ����TR�i���l/���l/�O���I�l�j
Public Sub Apply_Correct_Fields()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then last = 31   '30�����z��

    Application.ScreenUpdating = False

    ' @RssMarket �� RssMarket �ɏC��
    D.UsedRange.Replace What:="@RssMarket", Replacement:="RssMarket", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' ��{���ځi�ĕ~�݁j
            D.Cells(r, "B").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""��������""),"""")"
            D.Cells(r, "C").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),NA())"
            D.Cells(r, "D").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�n�l""),NA())"
            D.Cells(r, "E").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""���l""),NA())"
            D.Cells(r, "F").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""���l""),NA())"
            D.Cells(r, "G").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�o����""),0)"

            ' H: �o�������d���ρi��VWAP�j
            D.Cells(r, "H").Formula2 = _
                "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�o�������d����""),NA())"

            ' I: ����TR�iATR�������̂ő�ցj
            D.Cells(r, "I").Formula2 = _
                "=IFERROR(MAX($E" & r & "-$F" & r & ",ABS($E" & r & "-RssMarket(TEXT($A" & r & ",""0""),""�O���I�l""))," & _
                "ABS($F" & r & "-RssMarket(TEXT($A" & r & ",""0""),""�O���I�l""))),NA())"

            ' ����J, ���mK, ����L�iI����ɂ��̂܂܁j
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"

            ' �Q�l��iU..Z ���ĕ~�݁j
            D.Cells(r, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""20�����ϔ������""),0)"
            D.Cells(r, "V").Formula2 = _
                "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z"")-RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z""))/RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),1)"
            D.Cells(r, "W").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"  'ATR�~���i �� TR�~���i
            D.Cells(r, "X").Formula2 = "=$C" & r
            D.Cells(r, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s��敪""),"""")"
            D.Cells(r, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r & ")),ISNUMBER(SEARCH(""REIT"",Y" & r & "))),1,0)"
        End If
    Next

    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "H=�o�������d���� / I=����TR �ɏC���ς݁B", vbInformation
End Sub


