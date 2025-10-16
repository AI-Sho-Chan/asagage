Attribute VB_Name = "ASG_Optimize"
Option Explicit

' グリッドサーチ：利確係数 k_tp、損切係数 k_sl を0.5?3.0の範囲で探索
Public Sub GridSearch_TP_SL()
    Dim D As Worksheet: Set D = Worksheets("Dashboard")
    Dim S As Worksheet: Set S = Worksheets("Settings")
    Dim L As Worksheet
    On Error Resume Next
    Set L = Worksheets("PatchLog")
    On Error GoTo 0
    If L Is Nothing Then
        Set L = Worksheets.Add
        L.Name = "PatchLog"
    End If

    Dim last As Long: last = D.Cells(D.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then Exit Sub

    Dim k_tp As Double, k_sl As Double
    Dim bestScore As Double: bestScore = -1E+30
    Dim best_tp As Double, best_sl As Double

    Dim rowOut As Long: rowOut = 2
    L.Cells(1, 1).Value = "k_tp"
    L.Cells(1, 2).Value = "k_sl"
    L.Cells(1, 3).Value = "平均純利確"
    L.Cells(1, 4).Value = "有効件数"

    For k_tp = 0.5 To 3 Step 0.25
        For k_sl = 0.5 To 3 Step 0.25
            ' Settingsに適用
            S.Range("B22").Value = k_tp
            S.Range("B23").Value = k_sl
            Application.CalculateFull

            ' 純利確R列を評価
            Dim sumR As Double, cnt As Long, r As Long
            sumR = 0: cnt = 0
            For r = 2 To last
                If IsNumeric(D.Cells(r, "R").Value) Then
                    If D.Cells(r, "R").Value <> 0 Then
                        sumR = sumR + D.Cells(r, "R").Value
                        cnt = cnt + 1
                    End If
                End If
            Next

            If cnt > 0 Then
                Dim avgR As Double
                avgR = sumR / cnt
                L.Cells(rowOut, 1).Value = k_tp
                L.Cells(rowOut, 2).Value = k_sl
                L.Cells(rowOut, 3).Value = avgR
                L.Cells(rowOut, 4).Value = cnt
                rowOut = rowOut + 1

                If avgR > bestScore Then
                    bestScore = avgR: best_tp = k_tp: best_sl = k_sl
                End If
            End If
        Next
    Next

    MsgBox "探索終了。最良 k_tp=" & best_tp & " k_sl=" & best_sl & " 平均純利確=" & bestScore, vbInformation
End Sub

