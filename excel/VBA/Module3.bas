Attribute VB_Name = "Module3"
Option Explicit

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
        nextTick = Now + TimeSerial(0, 0, 2) ' 2�b�Ԋu
        Application.OnTime nextTick, "CheckSignals"
    End If
End Sub

Public Sub CheckSignals()
    On Error GoTo F
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Dashboard")
    Dim thr As Double: thr = ThisWorkbook.Sheets("Settings").Range("B3").Value
    Dim r As Long, hit As Boolean

    ' �O��̋������N���A
    sh.Range("A2:L11").Interior.ColorIndex = xlNone

    For r = 2 To 11
        If sh.Cells(r, "J").Value <> "" Then
            If Abs(sh.Cells(r, "J").Value) >= thr Then
                hit = True
                Beep
                sh.Rows(r).Interior.Color = RGB(255, 255, 160) ' �s��W���n�C���C�g
            End If
        End If
    Next r
GoTo X
F: ' ���������Ă��~�܂�Ȃ�
X:
    ScheduleNext
End Sub

'=== watchlist.txt ���������[�h ===
Private lastM As Date

Public Sub StartWatchlistAutoReload()
    lastM = 0
    WatchlistTick
End Sub

Private Sub WatchlistTick()
    On Error Resume Next
    Dim p As String: p = ThisWorkbook.Sheets("Settings").Range("B2").Value
    If Len(p) > 0 Then
        Dim m As Date: m = FileDateTime(p)
        If m > 0 And m <> lastM Then
            lastM = m
            LoadWatchlist        ' modAsagake.bas �̓Ǎ��֐����Ă�
        End If
    End If
    Application.OnTime Now + TimeSerial(0, 0, 5), "WatchlistTick" ' 5�b��
End Sub

