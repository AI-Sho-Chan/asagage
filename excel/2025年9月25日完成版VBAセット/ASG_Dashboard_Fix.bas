Attribute VB_Name = "ASG_Dashboard_Fix"
Option Explicit

Private Function EV(ByVal S As String) As Variant
    On Error GoTo e: EV = Application.Evaluate(S): Exit Function
e: EV = CVErr(xlErrNA)
End Function

'? 推定ティック（簡易）
Private Function TickSize(ByVal px As Double) As Double
    If px < 1000# Then TickSize = 1# Else TickSize = 1#
End Function

'? 予定発注数量（予算/価格/単元株）
Private Function PlannedQty(ByVal px As Double) As Double
    Dim budget#, lot#
    budget = Worksheets("Settings").Range("B35").Value
    lot = Worksheets("Settings").Range("B36").Value
    If px <= 0 Or lot <= 0 Then PlannedQty = 0: Exit Function
    PlannedQty = Int((budget / (px * lot))) * lot
End Function

'? スリッページ推計（場中=板、場外=日中量回帰）
Private Function SlipEst(ByVal code As String, ByVal side As String, ByVal px As Double, ByVal qty As Double) As Double
    Dim ask#, bid#, aSz#, bSz#, spread#, adv#
    ask = NzNum(EV("RssMarket(""" & code & """,""最良売気配値"")"))
    bid = NzNum(EV("RssMarket(""" & code & """,""最良買気配値"")"))
    aSz = NzNum(EV("RssMarket(""" & code & """,""最良売気配数量"")"))
    bSz = NzNum(EV("RssMarket(""" & code & """,""最良買気配数量"")"))
    spread = Application.Max(0, ask - bid)
    adv = NzNum(EV("RssMarket(""" & code & """,""出来高"")")) * px
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

    ' 見出しを確定（齟齬解消）
    D.Range("O1").Value = "ネット利確幅(円/株)" 'K ? (Q+R)
    D.Range("P1").Value = "予定発注数量[株]"
    D.Range("Q1").Value = "予想スリッページ:エントリー[円/株]"
    D.Range("R1").Value = "予想スリッページ:決済[円/株]"
    D.Range("S1").Value = "最終判定"
    D.Range("T1").Value = "売買代金"
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

    ' @ を除去
    D.UsedRange.Replace "@RssMarket", "RssMarket", xlPart

    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)
    Application.ScreenUpdating = False

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            code = Format$(D.Cells(r, "A").Value, "0")

            ' H: VWAP
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""出来高加重平均""),NA())"
            ' I: TR
            D.Cells(r, "I").Formula2 = "=IFERROR(MAX($E" & r & "-$F" & r & _
                ",ABS($E" & r & "-RssMarket(TEXT($A" & r & ",""0""),""前日終値""))," & _
                "ABS($F" & r & "-RssMarket(TEXT($A" & r & ",""0""),""前日終値""))),NA())"
            ' J: 乖離
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            ' K/L: 実現的な幅（TRと高安で上限制約）
            D.Cells(r, "K").Formula2 = "=IFERROR(MIN(Settings!$B$22*$I" & r & ",0.8*($E" & r & "-$F" & r & ")),NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR(MIN(Settings!$B$23*$I" & r & ",1.0*($E" & r & "-$F" & r & ")),NA())"

            ' W=価格
            D.Cells(r, "W").Formula2 = "=$C" & r
            ' X=市場区分
            D.Cells(r, "X").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場部名称""),IFERROR(RssMarket(TEXT($A" & r & ",""0""),""市場名称""),""""))"
            ' Y=除外
            D.Cells(r, "Y").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",X" & r & ")),ISNUMBER(SEARCH(""REIT"",X" & r & "))),1,0)"
            ' T,U,V
            D.Cells(r, "T").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""売買代金""),0)"
            D.Cells(r, "U").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r & ",""0""),""最良売気配値"")-RssMarket(TEXT($A" & r & ",""0""),""最良買気配値""))/RssMarket(TEXT($A" & r & ",""0""),""現在値""),0)"
            D.Cells(r, "V").Formula2 = "=IFERROR($I" & r & "*$C" & r & ",0)"

            ' P=予定発注数量
            D.Cells(r, "P").Value = PlannedQty(NzNum(D.Cells(r, "C").Value))

            ' Q/R スリッページ（値で埋める）
            D.Cells(r, "Q").Value = SlipEst(code, IIf(D.Cells(r, "J").Value < 0, "BUY", "SELL"), NzNum(D.Cells(r, "C").Value), NzNum(D.Cells(r, "P").Value))
            D.Cells(r, "R").Value = SlipEst(code, IIf(D.Cells(r, "J").Value < 0, "SELL", "BUY"), NzNum(D.Cells(r, "C").Value), NzNum(D.Cells(r, "P").Value))

            ' O=ネット利確幅
            D.Cells(r, "O").Formula2 = "=IFERROR($K" & r & "-($Q" & r & "+$R" & r & "),NA())"

            ' M=シグナル（J閾値）
            If IsNumeric(D.Cells(r, "J").Value) Then
                If Abs(D.Cells(r, "J").Value) >= j_th Then
                    D.Cells(r, "M").Value = IIf(D.Cells(r, "J").Value < 0, "エントリーシグナル", "ショートシグナル")
                Else
                    D.Cells(r, "M").ClearContents
                End If
            Else
                D.Cells(r, "M").ClearContents
            End If

            ' S=最終判定（ネット利確と条件）
            D.Cells(r, "S").Formula2 = "=IF(AND($O" & r & ">=" & minS & ",$M" & r & "<>"""",$AD" & r & "),IF($J" & r & "<0,""GO LONG"",""GO SHORT""),""SKIP"")"
        End If
    Next

    ' Z..AD（z標準化と条件OK）
    D.Range("Z2:Z31").Formula2 = "=IFERROR((T2-AVERAGE($T$2:$T$31))/STDEV.P($T$2:$T$31),0)"
    D.Range("AA2:AA31").Formula2 = "=IFERROR((V2-AVERAGE($V$2:$V$31))/STDEV.P($V$2:$V$31),0)"
    D.Range("AB2:AB31").Formula2 = "=IFERROR((U2-AVERAGE($U$2:$U$31))/STDEV.P($U$2:$U$31),0)"
    D.Range("AC2:AC31").Formula2 = "=0.6*Z2+0.5*AA2-0.7*AB2"
    D.Range("AD2:AD31").Formula2 = "=AND($W2>=500,$W2<=15000,$U2<=0.0025,$I2>=" & S.Range("B26").Value & ",$Y2=0)"

    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "Dashboard修正完了：8306含め非現実幅→現実的幅へ、J閾値連動、ネット利確(O)、最終判定(S)。", vbInformation
End Sub


