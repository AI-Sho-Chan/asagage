Attribute VB_Name = "ASG_Correctors"
Option Explicit

Public Sub Fix_U_V_W_X_Y()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)

    ' @ を全除去
    D.UsedRange.Replace What:="@RssMarket", Replacement:="RssMarket", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' U: 売買代金（週平均は未提供 → 当日値を採用）
            D.Cells(r, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""売買代金""),0)"
            ' V: スプレッド率＝(最良売気配値?最良買気配値)/現在値
            D.Cells(r, "V").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""最良売気配値"")-RssMarket(TEXT($A" & r & ",""0""),""最良買気配値""))/RssMarket(TEXT($A" & r & ",""0""),""現在値""),0)"
            ' W: TR×価格（IがTR=円、Xは価格）
            D.Cells(r, "W").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"
            D.Cells(r, "X").Formula2 = "=$C" & r
            ' Y: 市場部名称（無ければ市場名称に変える）
            D.Cells(r, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場部名称""),IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場名称""),""""))"
            ' Z: ETF/REIT 除外
            D.Cells(r, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r & ")),ISNUMBER(SEARCH(""REIT"",Y" & r & "))),1,0)"
        End If
    Next
    Application.CalculateFull
    MsgBox "U/V/W/X/Y を正しい関数で再敷設。", vbInformation
End Sub

Public Sub Fix_P_Q_R_and_Settings()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim S As Worksheet: Set S = Worksheets("Settings")
    Dim last As Long, r As Long
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)

    ' 係数見直し（IはTR=円なので 150→1.0）
    S.Range("B22").Value = 1     '利確係数
    S.Range("B23").Value = 1     '損切係数

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' P: 予想スリッページ(ENTRY)
            D.Cells(r, "P").Formula2 = _
              "=IF($M" & r & "="""","""",W_ENTRY_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),$C" & r & ",1,1,1,Settings!$B$27))"
            ' Q: 予想スリッページ(EXIT)
            D.Cells(r, "Q").Formula2 = _
              "=IF($M" & r & "="""","""",W_EXIT_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""SELL"",""BUY""),$C" & r & ",W_QTY_BY_BUDGET_CLIPPED($C" & r & ",1,1,$A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),1),1,Settings!$B$27))"
            ' R: 純利確幅
            D.Cells(r, "R").Formula2 = "=IF($M" & r & "="""","""",MAX(0,$O" & r & "-$P" & r & "-$Q" & r & "))"
        End If
    Next
    Application.CalculateFull
    MsgBox "P/Q/R を修正。Settings!B22/B23 を 1.0 に変更。", vbInformation
End Sub


