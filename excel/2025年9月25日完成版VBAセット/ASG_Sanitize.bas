Attribute VB_Name = "ASG_Sanitize"
Option Explicit

' UNIVERSE_EXTRA のコードを4桁数字に正規化し、
' Dashboard の @RssMarket を RssMarket に置換、式の列ズレも復元する
Public Sub Sanitize_Codes_And_Formulas()
    Dim U As Worksheet, D As Worksheet, r As Long, last As Long, v As String, t As String

    Set U = Worksheets("UNIVERSE_EXTRA")
    Set D = Worksheets("Dashboard")

    ' 1) コード正規化（A列 4桁数字以外は削除）
    last = U.Cells(U.Rows.Count, "A").End(xlUp).Row
    For r = 2 To last
        v = CStr(U.Cells(r, "A").Value)
        t = Digits4(v)
        If Len(t) = 4 Then
            U.Cells(r, "A").Value = t
        Else
            U.Cells(r, "A").ClearContents
        End If
    Next

    ' 2) Dashboard の @RssMarket を一括修正
    Call ReplaceInRange(D.Range("B2:AE3000"), "@RssMarket", "RssMarket")

    ' 3) 必須列の式を再敷設（列ズレ修復）
    Dim lastRow As Long: lastRow = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 300

    Dim r2 As Long
    For r2 = 2 To lastRow
        D.Cells(r2, "B").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""銘柄名称""),"""")"
        D.Cells(r2, "C").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""現在値""),NA())"
        D.Cells(r2, "D").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""始値""),NA())"
        D.Cells(r2, "E").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""高値""),NA())"
        D.Cells(r2, "F").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""安値""),NA())"
        D.Cells(r2, "G").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""出来高""),0)"
        ' H/I は RssMarket を先に入れておく（Barsが無くても埋まる）
        D.Cells(r2, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""当日VWAP""),NA())"
        D.Cells(r2, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""ATR(5)""),NA())"
        D.Cells(r2, "J").Formula2 = "=IFERROR(($C" & r2 & "-$H" & r2 & ")/$I" & r2 & ",NA())"
        D.Cells(r2, "K").Formula2 = "=IFERROR($I" & r2 & "*Settings!$B$22,NA())"
        D.Cells(r2, "L").Formula2 = "=IFERROR($I" & r2 & "*Settings!$B$23,NA())"
        D.Cells(r2, "O").Formula2 = "=IFERROR($C" & r2 & "-$L" & r2 & ",NA())"
        D.Cells(r2, "U").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""20日平均売買代金""),0)"
        D.Cells(r2, "V").Formula2 = "=IFERROR((RssMarket(TEXT($A" & r2 & ",""0""),""最良売気配"")-RssMarket(TEXT($A" & r2 & ",""0""),""最良買気配""))/RssMarket(TEXT($A" & r2 & ",""0""),""現在値""),1)"
        D.Cells(r2, "W").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""ATR(5)"")*RssMarket(TEXT($A" & r2 & ",""0""),""現在値""),0)"
        D.Cells(r2, "X").Formula2 = "=$C" & r2
        D.Cells(r2, "Y").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r2 & ",""0""),""市場区分""),"""")"
        D.Cells(r2, "Z").Formula2 = "=IF(OR(ISNUMBER(SEARCH(""ETF"",Y" & r2 & ")),ISNUMBER(SEARCH(""REIT"",Y" & r2 & "))),1,0)"
    Next

    Application.CalculateFull
    MsgBox "Sanitize 完了：UNIVERSE_EXTRAのコード正規化＆Dashboard式修復", vbInformation
End Sub

' ---- helpers ----
Private Function Digits4(ByVal S As String) As String
    Dim i As Long, cH As String, t As String
    For i = 1 To Len(S)
        cH = Mid$(S, i, 1)
        If cH Like "#" Then t = t & cH
        If Len(t) = 4 Then Exit For
    Next
    Digits4 = t
End Function

Private Sub ReplaceInRange(ByVal rg As Range, ByVal findText As String, ByVal repText As String)
    With rg
        .Replace What:=findText, Replacement:=repText, LookAt:=xlPart, _
                 SearchOrder:=xlByRows, MatchCase:=False
    End With
End Sub


