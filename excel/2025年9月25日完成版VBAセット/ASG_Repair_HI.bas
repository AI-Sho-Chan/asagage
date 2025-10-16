Attribute VB_Name = "ASG_Repair_HI"
Option Explicit

Public Sub Repair_HI_And_Test()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long, code As String
    last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then last = 31

    Application.ScreenUpdating = False

    ' 1) H/I の書式と内容を初期化（文字列→数式に戻す）
    With D.Range("H2:I" & last)
        .NumberFormat = "General"
        .Replace What:="'=", Replacement:="=", LookAt:=xlPart
        .Replace What:="@RssMarket", Replacement:="RssMarket", LookAt:=xlPart
        .ClearContents
    End With

    ' 2) H/I の数式を再敷設（RssMarket 直）
    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""当日VWAP""),NA())"
            D.Cells(r, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""ATR(5)""),NA())"
        End If
    Next

    ' 3) J/K/L も再計算式を保証
    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"
        End If
    Next

    Application.CalculateFull
    Application.ScreenUpdating = True

    ' 4) RSSアドイン疎通テスト（A2のコードで1回だけ）
    code = Trim$(CStr(D.Cells(2, "A").Value))
    If Len(code) = 0 Then
        MsgBox "A2 が空です。まずコードを入れてください。", vbExclamation: Exit Sub
    End If
    Dim t1 As Variant, t2 As Variant
    t1 = Application.Evaluate("RssMarket(""" & code & """,""当日VWAP"")")
    t2 = Application.Evaluate("RssMarket(""" & code & """,""ATR(5)"")")

    If IsError(t1) Or IsError(t2) Then
        MsgBox "RSS未応答またはフィールド名不一致。「=RssMarket(""" & code & """,""当日VWAP"")」を任意セルに直打ちして値が返るか確認。返らなければMarketSpeed IIのアドインを有効化し、Excel再起動。", vbExclamation
    Else
        MsgBox "H/I 式を再敷設しました。テスト値: VWAP=" & t1 & " / ATR5=" & t2, vbInformation
    End If
End Sub

