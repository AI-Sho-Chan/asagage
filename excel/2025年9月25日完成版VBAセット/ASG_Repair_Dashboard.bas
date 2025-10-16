Attribute VB_Name = "ASG_Repair_Dashboard"
Option Explicit

' A2:A31 �ɃR�[�h������O��ŁAB�`G/H/I/J/K/L/O/U�`Z �����ňꊇ�~��
Public Sub Fill_Dashboard_30()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim r As Long
    Application.ScreenUpdating = False
    For r = 2 To 31
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "B").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""��������""),"""")"
            D.Cells(r, "C").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),NA())"
            D.Cells(r, "D").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�n�l""),NA())"
            D.Cells(r, "E").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""���l""),NA())"
            D.Cells(r, "F").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""���l""),NA())"
            D.Cells(r, "G").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�o����""),0)"
            ' H/I �� RssMarket ���iBars�͎g��Ȃ��j
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""����VWAP""),NA())"
            D.Cells(r, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""ATR(5)""),NA())"
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"
            D.Cells(r, "O").Formula2 = "=IFERROR($C" & r & "-$L" & r & ",NA())"
            D.Cells(r, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""20�����ϔ������""),0)"
            D.Cells(r, "V").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z"")-RssMarket(TEXT($A" & r & ",""0""),""�ŗǔ��C�z""))/RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),1)"
            D.Cells(r, "W").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""ATR(5)"")*RssMarket(TEXT($A" & r & ",""0""),""���ݒl""),0)"
            D.Cells(r, "X").Formula2 = "=$C" & r
            D.Cells(r, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""�s��敪""),"""")"
            D.Cells(r, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r & ")),ISNUMBER(SEARCH(""REIT"",Y" & r & "))),1,0)"
        End If
    Next
    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "Dashboard formulas (H/I�܂�) ��A2:A31�ɕ~�݂��܂����B", vbInformation
End Sub

