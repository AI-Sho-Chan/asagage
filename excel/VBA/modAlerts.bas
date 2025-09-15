Attribute VB_Name = "modAlerts"
Option Explicit

'=== アラート（2秒ごとにチェック）===
Private nextTick As Date
Private running As Boolean

Public Sub StartAlerts()
    running = True
    ScheduleNext
End Sub

Public Sub StopAlerts()
    On Error Resume Next
    running = False
    Application.OnTime earliesttime:=nextTick, procedure:="CheckSignals", schedule:=False
End Sub

Private Sub ScheduleNext()
    If running Then
        nextTick = Now + TimeSerial(0, 0, 2)
        Application.OnTime nextTick, "CheckSignals"
    End If
End Sub

Public Sub CheckSignals()
    On Error GoTo f
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Dashboard")
    Dim thr As Double: thr = ThisWorkbook.Sheets("Settings").Range("B3").Value
    Dim lastRow As Long: lastRow = Application.Max(2, sh.Cells(sh.Rows.Count, "A").End(xlUp).Row)
    Dim r As Long

    sh.Range("A2:L" & lastRow).Interior.ColorIndex = xlNone
    For r = 2 To lastRow
        If sh.Cells(r, "A").Value <> "" And sh.Cells(r, "J").Value <> "" Then
            If Abs(sh.Cells(r, "J").Value) >= thr Then
                Beep
                sh.Rows(r).Interior.Color = RGB(255, 255, 160)
            End If
        End If
    Next r
f:
    ScheduleNext
End Sub

