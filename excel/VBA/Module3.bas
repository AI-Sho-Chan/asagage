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
        nextTick = Now + TimeSerial(0, 0, 2) ' 2秒間隔
        Application.OnTime nextTick, "CheckSignals"
    End If
End Sub

Public Sub CheckSignals()
    On Error GoTo F
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Dashboard")
    Dim thr As Double: thr = ThisWorkbook.Sheets("Settings").Range("B3").Value
    Dim r As Long, hit As Boolean

    ' 前回の強調をクリア
    sh.Range("A2:L11").Interior.ColorIndex = xlNone

    For r = 2 To 11
        If sh.Cells(r, "J").Value <> "" Then
            If Abs(sh.Cells(r, "J").Value) >= thr Then
                hit = True
                Beep
                sh.Rows(r).Interior.Color = RGB(255, 255, 160) ' 行を淡くハイライト
            End If
        End If
    Next r
GoTo X
F: ' 何かあっても止まらない
X:
    ScheduleNext
End Sub

'=== watchlist.txt 自動リロード ===
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
            LoadWatchlist        ' modAsagake.bas の読込関数を呼ぶ
        End If
    End If
    Application.OnTime Now + TimeSerial(0, 0, 5), "WatchlistTick" ' 5秒毎
End Sub

