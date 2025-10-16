Attribute VB_Name = "ASG_Signal"
Option Explicit
' ルール：Jの符号で方向決定、R(純利確) > Settings!B24 かつ AE=TRUE でGO
Public Sub Make_Signal()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim last As Long, r As Long
    last = 31
    Application.ScreenUpdating = False
    For r = 2 To last
        If Len(Trim$(D.Cells(r, "A").Value)) > 0 Then
            ' シグナル（M）
            If IsNumeric(D.Cells(r, "J").Value) Then
                If D.Cells(r, "J").Value < 0 Then
                    D.Cells(r, "M").Value = "ショートシグナル"
                Else
                    D.Cells(r, "M").Value = "エントリーシグナル"
                End If
            Else
                D.Cells(r, "M").ClearContents
            End If
            ' タイムスタンプ（N）
            D.Cells(r, "N").Value = Now
            ' 予想スリッページ類（P,Q,R）は式または別途UDFで運用。ここでは未変更
            ' 最終判定（S）
            D.Cells(r, "S").Formula2 = _
                "=IF(AND($R" & r & ">=Settings!$B$24,$AE" & r & "=TRUE),IF($J" & r & "<0,""GO SHORT"",""GO LONG""),""SKIP"")"
        End If
    Next
    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "M/N/S を更新しました。", vbInformation
End Sub


