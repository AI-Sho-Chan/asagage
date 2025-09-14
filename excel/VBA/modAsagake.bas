Option Explicit

Public Sub Auto_Open()
    LoadWatchlist
End Sub

Public Sub LoadWatchlist()
    On Error GoTo EH
    Dim p As String: p = ThisWorkbook.Sheets("Settings").Range("B2").Value
    Dim f As Integer: f = FreeFile
    Dim line As String, r As Long: r = 2
    Open p For Input As #f
    With Sheets("Dashboard")
        .Range("A2:A100").ClearContents
        Do While Not EOF(f)
            Line Input #f, line
            line = Trim(line)
            If Len(line) > 0 Then
                .Cells(r, 1).Value = line
                r = r + 1
            End If
        Loop
    End With
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    Close #f
    MsgBox "watchlist.txt を読み込めません: " & p, vbExclamation
End Sub
