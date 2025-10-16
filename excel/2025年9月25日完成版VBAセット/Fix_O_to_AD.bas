Attribute VB_Name = "Fix_O_to_AD"
Option Explicit

Public Sub Fix_O_to_AD()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)

    ' 見出し統一
    D.Range("O1").Value = "利確幅(円)"
    D.Range("P1").Value = "約定単位"
    D.Range("Q1").Value = "予想スリッページ:エントリー"
    D.Range("R1").Value = "予想スリッページ:決済"
    D.Range("S1").Value = "最終判定"
    D.Range("T1").Value = "売買代金"            '※20日平均は未提供 → 当日値を採用
    D.Range("U1").Value = "スプレッド率"
    D.Range("V1").Value = "TR×価格"
    D.Range("W1").Value = "価格"
    D.Range("X1").Value = "市場区分"
    D.Range("Y1").Value = "除外フラグ"
    D.Range("Z1").Value = "z_流動性"
    D.Range("AA1").Value = "z_ボラ"
    D.Range("AB1").Value = "z_スプレッド"
    D.Range("AC1").Value = "総合S"
    D.Range("AD1").Value = "条件OK"

    ' @RssMarket → RssMarket
    D.UsedRange.Replace What:="@RssMarket", Replacement:="RssMarket", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    Application.ScreenUpdating = False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' O: 利確幅
            D.Cells(r, "O").Formula2 = "=IFERROR($C" & r & "-$L" & r & ",NA())"
            ' P: 約定単位（単元情報が未提供のため既定=100）
            D.Cells(r, "P").Formula2 = "=IF($M" & r & "="""","""",100)"
            ' Q/R: スリッページ
            D.Cells(r, "Q").Formula2 = _
              "=IF($M" & r & "="""","""",W_ENTRY_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),$C" & r & ",Settings!$B$22,Settings!$B$23,Settings!$B$31,Settings!$B$27))"
            D.Cells(r, "R").Formula2 = _
              "=IF($M" & r & "="""","""",W_EXIT_SLIP_CAP($A" & r & ",IF($J" & r & "<0,""SELL"",""BUY""),$C" & r & ",W_QTY_BY_BUDGET_CLIPPED($C" & r & ",Settings!$B$22,Settings!$B$23,$A" & r & ",IF($J" & r & "<0,""BUY"",""SELL""),Settings!$B$31),Settings!$B$31,Settings!$B$27))"
            ' S: 最終判定
            D.Cells(r, "S").Formula2 = "=IF(AND($R" & r & ">=Settings!$B$24,$AD" & r & "=TRUE),IF($J" & r & "<0,""GO SHORT"",""GO LONG""),""SKIP"")"
            ' T: 売買代金
            D.Cells(r, "T").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""売買代金""),0)"
            ' U: スプレッド率
            D.Cells(r, "U").Formula2 = _
              "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""最良売気配値"")-RssMarket(TEXT($A" & r & ",""0""),""最良買気配値""))/RssMarket(TEXT($A" & r & ",""0""),""現在値""),0)"
            ' V: TR×価格
            D.Cells(r, "V").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"
            ' W: 価格
            D.Cells(r, "W").Formula2 = "=$C" & r
            ' X: 市場区分
            D.Cells(r, "X").Formula2 = _
              "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場部名称""),IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場名称""),""""))"
            ' Y: 除外フラグ（X参照で循環回避）
            D.Cells(r, "Y").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",X" & r & ")),ISNUMBER(SEARCH(""REIT"",X" & r & "))),1,0)"

        End If
    Next

    ' Z,AA,AB の zスコア & AC/AD
    ' 範囲は2:31想定。ゼロ割回避。
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
    MsgBox "O?AD の見出しと式を修正しました。", vbInformation
End Sub


