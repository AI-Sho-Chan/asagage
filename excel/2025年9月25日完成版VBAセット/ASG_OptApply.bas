Attribute VB_Name = "ASG_OptApply"
Option Explicit

Public Sub Optimize_And_Apply()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim s As Worksheet: Set s = Worksheets("Settings")
    Dim L As Worksheet: On Error Resume Next: Set L = Worksheets("PatchLog"): On Error GoTo 0
    If L Is Nothing Then Set L = Worksheets.Add(After:=D): L.Name = "PatchLog"

    Dim last&, r&, rowOut&, k_tp#, k_sl#, best_tp#, best_sl#, bestScore#
    last = Application.Max(31, D.Cells(D.Rows.Count, "A").End(xlUp).Row)
    L.Range("A1:D1").Value = Array("k_tp", "k_sl", "平均ネット利確(O)", "有効件数"): rowOut = 2
    bestScore = -1E+30

    For k_tp = 0.5 To 3 Step 0.25
        For k_sl = 0.5 To 3 Step 0.25
            s.Range("B22").Value = k_tp
            s.Range("B23").Value = k_sl
            Application.CalculateFull

            Dim sumO#, cnt&
            For r = 2 To last
                If IsNumeric(D.Cells(r, "O").Value) Then sumO = sumO + D.Cells(r, "O").Value: cnt = cnt + 1
            Next
            If cnt > 0 Then
                Dim avgO#: avgO = sumO / cnt
                L.Cells(rowOut, 1).Resize(1, 4).Value = Array(k_tp, k_sl, avgO, cnt): rowOut = rowOut + 1
                If avgO > bestScore Then bestScore = avgO: best_tp = k_tp: best_sl = k_sl
            End If
        Next
    Next

    s.Range("B22").Value = best_tp
    s.Range("B23").Value = best_sl
    Application.CalculateFull
    MsgBox "最適化完了: k_tp=" & best_tp & "  k_sl=" & best_sl & "  平均O=" & Format(bestScore, "0.00"), vbInformation
End Sub

