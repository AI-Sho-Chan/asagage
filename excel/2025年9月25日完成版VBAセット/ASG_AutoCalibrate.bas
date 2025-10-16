Attribute VB_Name = "ASG_AutoCalibrate"
Option Explicit

' 候補のうち実際に値が返る項目名をランタイム同定
Private Function ResolveField(ByVal code As String, ByRef candidates() As String) As String
    Dim i As Long, v As Variant
    For i = LBound(candidates) To UBound(candidates)
        v = Application.Evaluate("RssMarket(""" & code & """,""" & candidates(i) & """)")
        If Not IsError(v) Then
            If IsNumeric(v) Or Len(CStr(v)) > 0 Then
                ResolveField = candidates(i)
                Exit Function
            End If
        End If
    Next
    ResolveField = ""
End Function

' H/I の式を “正しい項目名” で敷設し直す（A2:A31 が対象）
Public Sub Calibrate_And_Fill_HI()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim code As String
    code = Trim$(CStr(D.Cells(2, "A").Value))
    If Len(code) = 0 Then MsgBox "A2 が空です。コードを入れてください。", vbExclamation: Exit Sub

    ' 候補リスト（実機で返るものを探索）
    Dim vwapCands() As String, atrCands() As String
    vwapCands = Split("VWAP,当日VWAP,出来高加重平均,当日出来高加重平均,当日出来高加重平均価格,加重平均", ",")
    atrCands = Split("ATR(5),ATR 5,アベレージ・トゥルー・レンジ(5),ATR,ATR(14),ATR % (14) 1日,ATR％(14) 1日", ",")

    Dim vwapKey As String, atrKey As String
    vwapKey = ResolveField(code, vwapCands)
    atrKey = ResolveField(code, atrCands)

    If vwapKey = "" Then
        MsgBox "VWAP系の項目名が見つかりません。手動で正名を教えてください。", vbCritical
        Exit Sub
    End If
    If atrKey = "" Then
        MsgBox "ATR系の項目名が見つかりません。手動で正名を教えてください。", vbCritical
        Exit Sub
    End If

    ' 設置
    Dim last As Long, r As Long
    last = 31  ' 30銘柄固定運用
    Application.ScreenUpdating = False

    ' H/I をクリアして確実に数式化
    With D.Range("H2:I" & last)
        .NumberFormat = "General"
        .Replace What:="'=", Replacement:="=", LookAt:=xlPart
        .Replace What:="@RssMarket", Replacement:="RssMarket", LookAt:=xlPart
        .ClearContents
    End With

    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            D.Cells(r, "H").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""" & vwapKey & """),NA())"
            D.Cells(r, "I").Formula2 = "=IFERROR(RssMarket(TEXT($A" & r & ",""0""),""" & atrKey & """),NA())"
            ' 依存列も再敷設
            D.Cells(r, "J").Formula2 = "=IFERROR(($C" & r & "-$H" & r & ")/$I" & r & ",NA())"
            D.Cells(r, "K").Formula2 = "=IFERROR($I" & r & "*Settings!$B$22,NA())"
            D.Cells(r, "L").Formula2 = "=IFERROR($I" & r & "*Settings!$B$23,NA())"
        End If
    Next

    Application.CalculateFull

    MsgBox "VWAP=""" & vwapKey & """ / ATR=""" & atrKey & """ で H/I を再敷設しました。", vbInformation
End Sub

