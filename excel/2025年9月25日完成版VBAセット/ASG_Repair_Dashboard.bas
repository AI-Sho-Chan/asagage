Attribute VB_Name = "ASG_Repair_Dashboard"
Option Explicit

' A2:A31 にコードがある前提で、B～G/H/I/J/K/L/O/U～Z を式で一括敷設
Public Sub Fill_Dashboard_30()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim r As Long
    Application.ScreenUpdating = False
    For r = 2 To 31
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "B").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""銘柄名称""),"""")"
            D.Cells(r, "C").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""現在値""),NA())"
            D.Cells(r, "D").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""始値""),NA())"
            D.Cells(r, "E").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""高値""),NA())"
            D.Cells(r, "F").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""安値""),NA())"
            D.Cells(r, "G").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""出来高""),0)"
            ' H/I は RssMarket 直（Barsは使わない）
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""当日VWAP""),NA())"
            D.Cells(r, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""ATR(5)""),NA())"
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"
            D.Cells(r, "O").Formula2 = "=IFERROR($C" & r & "-$L" & r & ",NA())"
            D.Cells(r, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""20日平均売買代金""),0)"
            D.Cells(r, "V").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""最良売気配"")-RssMarket(TEXT($A" & r & ",""0""),""最良買気配""))/RssMarket(TEXT($A" & r & ",""0""),""現在値""),1)"
            D.Cells(r, "W").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""ATR(5)"")*RssMarket(TEXT($A" & r & ",""0""),""現在値""),0)"
            D.Cells(r, "X").Formula2 = "=$C" & r
            D.Cells(r, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場区分""),"""")"
            D.Cells(r, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r & ")),ISNUMBER(SEARCH(""REIT"",Y" & r & "))),1,0)"
        End If
    Next
    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "Dashboard formulas (H/I含む) をA2:A31に敷設しました。", vbInformation
End Sub

