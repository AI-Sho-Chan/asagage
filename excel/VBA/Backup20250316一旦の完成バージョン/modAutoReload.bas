Attribute VB_Name = "modAutoReload"

Option Explicit

Private nextWL As Date

Public Sub LoadWatchlist()
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Dashboard")
    Dim p As String: p = CStr(ThisWorkbook.Sheets("Settings").Range("B2").Value)
    Dim ff As Integer, r As Long, line As String
    If Len(dir(p)) = 0 Then Exit Sub
    ws.Range("A2:A200").ClearContents
    ff = FreeFile: Open p For Input As #ff
    r = 2
    Do While Not EOF(ff) And r <= 21
        Line Input #ff, line
        line = Trim$(line)
        If Len(line) > 0 Then
            ws.Cells(r, 1).Value = CleanTicker(line)
            r = r + 1
        End If
    Loop
    Close #ff
    Exit Sub
EH:
    On Error Resume Next: Close #ff
End Sub

Public Sub StartWatchlistAutoReload()
    StopWatchlistAutoReload
    WatchlistTick
End Sub

Public Sub StopWatchlistAutoReload()
    On Error Resume Next
    If nextWL <> 0 Then Application.OnTime earliesttime:=nextWL, procedure:="WatchlistTick", schedule:=False
End Sub

Public Sub WatchlistTick()
    On Error Resume Next
    LoadWatchlist
    nextWL = Now + TimeSerial(0, 0, 5)
    Application.OnTime earliesttime:=nextWL, procedure:="WatchlistTick"
    'RecordTick          '※必要なら有効化（modRecorder）
    'PaperTradeTick      '※必要なら有効化（modPaperTrader）
End Sub

