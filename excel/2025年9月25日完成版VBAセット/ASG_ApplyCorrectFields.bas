Attribute VB_Name = "ASG_ApplyCorrectFields"
Option Explicit

' DashboardのH/Iと依存列を「正しい関数」に修正
' H = 出来高加重平均、I = 当日TR（高値/安値/前日終値）
Public Sub Apply_Correct_Fields()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then last = 31   '30銘柄想定

    Application.ScreenUpdating = False

    ' @RssMarket を RssMarket に修正
    D.UsedRange.Replace What:="@RssMarket", Replacement:="RssMarket", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' 基本項目（再敷設）
            D.Cells(r, "B").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""銘柄名称""),"""")"
            D.Cells(r, "C").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""現在値""),NA())"
            D.Cells(r, "D").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""始値""),NA())"
            D.Cells(r, "E").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""高値""),NA())"
            D.Cells(r, "F").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""安値""),NA())"
            D.Cells(r, "G").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""出来高""),0)"

            ' H: 出来高加重平均（＝VWAP）
            D.Cells(r, "H").Formula2 = _
                "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""出来高加重平均""),NA())"

            ' I: 当日TR（ATRが無いので代替）
            D.Cells(r, "I").Formula2 = _
                "=IFERROR(MAX($E" & r & "-$F" & r & ",ABS($E" & r & "-RssMarket(TEXT($A" & r & ",""0""),""前日終値""))," & _
                "ABS($F" & r & "-RssMarket(TEXT($A" & r & ",""0""),""前日終値""))),NA())"

            ' 乖離J, 利確K, 損切L（Iを基準にそのまま）
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"

            ' 参考列（U..Z も再敷設）
            D.Cells(r, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""20日平均売買代金""),0)"
            D.Cells(r, "V").Formula2 = _
                "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""最良売気配"")-RssMarket(TEXT($A" & r & ",""0""),""最良買気配""))/RssMarket(TEXT($A" & r & ",""0""),""現在値""),1)"
            D.Cells(r, "W").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"  'ATR×価格 → TR×価格
            D.Cells(r, "X").Formula2 = "=$C" & r
            D.Cells(r, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場区分""),"""")"
            D.Cells(r, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r & ")),ISNUMBER(SEARCH(""REIT"",Y" & r & "))),1,0)"
        End If
    Next

    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "H=出来高加重平均 / I=当日TR に修正済み。", vbInformation
End Sub


