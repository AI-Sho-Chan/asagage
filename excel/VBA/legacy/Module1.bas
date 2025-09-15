Attribute VB_Name = "Module1"
Option Explicit

Public Sub Auto_Open()
    LoadWatchlist
End Sub

Public Sub LoadWatchlist()
    On Error GoTo EH
    Dim p As String: p = ThisWorkbook.Sheets("Settings").Range("B2").Value
    Dim F As Integer: F = FreeFile
    Dim line As String, r As Long: r = 2
    Open p For Input As #F
    With Sheets("Dashboard")
        .Range("A2:A100").ClearContents
        Do While Not EOF(F)
            Line Input #F, line
            line = Trim(line)
            If Len(line) > 0 Then
                .Cells(r, 1).Value = line
                r = r + 1
            End If
        Loop
    End With
    Close #F
    Exit Sub
EH:
    On Error Resume Next
    Close #F
    MsgBox "watchlist.txt を読み込めません: " & p, vbExclamation
End Sub
