Attribute VB_Name = "NightlyStatus"
Option Explicit

Private Const STATUS_FILE_REL As String = "\logs\nightly_status.txt"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const NEW_DASHBOARD_SHEET As String = "NewDashboard"
Private Const STATUS_HEADER_CELL As String = "AU2"
Private Const LOG_TABLE_NAME As String = "NightlyStatusLog"
Private Const REFRESH_SECONDS As Double = 10#

Private monitorActive As Boolean
Private scheduledTime As Date
Private lastUpdateStamp As String
Private lastLoggedStamp As String

Public Sub NotifyLaunch()
    lastUpdateStamp = vbNullString
    lastLoggedStamp = vbNullString
    RenderManualLaunchMessage
    ScheduleNext True
    RefreshNightlyStatus
End Sub

Public Sub InitializeOnOpen()
    lastLoggedStamp = vbNullString
    lastUpdateStamp = vbNullString
    Dim info As Object
    Set info = ReadStatusFile()
    If info Is Nothing Then
        RenderWaitingMessage "Nightly batch status file not found."
        ScheduleNext False
        Exit Sub
    End If
    lastUpdateStamp = NzDict(info, "updated", "")
    RenderStatus info
    AppendStatusLog info
    Dim state As String
    state = LCase$(NzDict(info, "state", ""))
    If state = "running" Or state = "" Then
        ScheduleNext False
    Else
        StopNightlyMonitor
    End If
End Sub

Public Sub RefreshNightlyStatus()
    On Error GoTo HandleFailure
    Dim info As Object
    Set info = ReadStatusFile()
    If info Is Nothing Then
        RenderWaitingMessage "Waiting for nightly batch status..."
        ScheduleNext False
        Exit Sub
    End If

    Dim updated As String
    updated = NzDict(info, "updated", "")
    If updated <> "" Then
        lastUpdateStamp = updated
    End If

    RenderStatus info
    AppendStatusLog info

    Dim state As String
    state = LCase$(NzDict(info, "state", ""))
    If state = "running" Or state = "" Then
        ScheduleNext False
    Else
        StopNightlyMonitor
    End If
    Exit Sub

HandleFailure:
    RenderWaitingMessage "Status update failed: " & Err.Description
    ScheduleNext False
End Sub

Public Sub StopNightlyMonitor()
    On Error Resume Next
    If monitorActive And scheduledTime > 0 Then
        Application.OnTime earliesttime:=scheduledTime, Procedure:="NightlyStatus.RefreshNightlyStatus", Schedule:=False
    End If
    monitorActive = False
    scheduledTime = 0
End Sub

Private Sub ScheduleNext(Optional ByVal immediate As Boolean = False)
    On Error Resume Next
    If monitorActive And scheduledTime > 0 Then
        Application.OnTime earliesttime:=scheduledTime, Procedure:="NightlyStatus.RefreshNightlyStatus", Schedule:=False
    End If
    On Error GoTo 0

    monitorActive = True
    Dim delay As Double
    If immediate Then
        delay = 1# / 86400#
    Else
        delay = REFRESH_SECONDS / 86400#
    End If
    scheduledTime = Now + delay
    Application.OnTime earliesttime:=scheduledTime, Procedure:="NightlyStatus.RefreshNightlyStatus"
End Sub

Private Function ReadStatusFile() As Object
    On Error GoTo CleanFail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = ThisWorkbook.path & STATUS_FILE_REL
    If Not fso.FileExists(path) Then Exit Function

    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1, False, -1)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Do Until ts.AtEndOfStream
        Dim line As String
        line = Trim$(ts.ReadLine)
        If Len(line) = 0 Then GoTo ContinueLoop
        Dim eqPos As Long
        eqPos = InStr(1, line, "=")
        If eqPos > 0 Then
            Dim key As String
            Dim value As String
            key = Trim$(Left$(line, eqPos - 1))
            value = Mid$(line, eqPos + 1)
            dict(key) = value
        End If
ContinueLoop:
    Loop
    ts.Close
    Set ReadStatusFile = dict
    Exit Function

CleanFail:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
End Function

Private Sub RenderStatus(ByVal info As Object)
    RenderStatusToSheet DASHBOARD_SHEET, info
    RenderStatusToSheet NEW_DASHBOARD_SHEET, info
End Sub

Private Sub RenderManualLaunchMessage()
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info("state") = "launching"
    info("message") = "Catch-up requested. Waiting for status updates..."
    info("started") = Format$(Now, "yyyy-mm-dd HH:NN:SS")
    info("updated") = info("started")
    RenderStatus info
End Sub

Private Sub RenderWaitingMessage(ByVal message As String)
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info("state") = "waiting"
    info("message") = message
    info("updated") = Format$(Now, "yyyy-mm-dd HH:NN:SS")
    RenderStatus info
End Sub

Private Sub RenderStatusToSheet(ByVal sheetName As String, ByVal info As Object)
    Dim ws As Worksheet
    Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Sub

    With ws
        .Range(STATUS_HEADER_CELL).Value = "Nightly Catch-Up Status"
        .Range("AU3").Value = "State"
        .Range("AV3").Value = NzDict(info, "state", "")
        .Range("AU4").Value = "Step"
        .Range("AV4").Value = NzDict(info, "step", "")
        .Range("AU5").Value = "Message"
        .Range("AV5").Value = NzDict(info, "message", "")
        .Range("AU6").Value = "Started"
        .Range("AV6").Value = NzDict(info, "started", "")
        .Range("AU7").Value = "Last Update"
        .Range("AV7").Value = NzDict(info, "updated", "")
        .Range("AU8").Value = "Elapsed"
        .Range("AV8").Value = FormatElapsedSeconds(NzDict(info, "elapsed_seconds", ""))
        .Range("AU9").Value = "Plans"
        .Range("AV9").Value = NzDict(info, "plans", "")
        .Range("AU10").Value = "Plan Counts"
        .Range("AV10").Value = NzDict(info, "plan_counts", "")
        .Range("AU11").Value = "Candidate Files"
        .Range("AV11").Value = NzDict(info, "candidate_files", "")
        .Range("AU12").Value = "Total Candidates"
        .Range("AV12").Value = NzDict(info, "total_candidates", "")
        .Range("AU13").Value = "Unique Tickers"
        .Range("AV13").Value = NzDict(info, "unique_tickers", "")
        .Range("AU14").Value = "Avg Forward Winrate"
        .Range("AV14").Value = NzDict(info, "avg_forward_winrate", "")
        .Range("AU15").Value = "Avg Forward PF"
        .Range("AV15").Value = NzDict(info, "avg_forward_pf", "")
        .Range("AU16").Value = "Avg Expected Return"
        .Range("AV16").Value = NzDict(info, "avg_expected_return", "")
        .Range("AU17").Value = "Avg Forward Trades"
        .Range("AV17").Value = NzDict(info, "avg_forward_trades", "")
        .Range("AU18").Value = "Output Path"
        .Range("AV18").Value = NzDict(info, "candidates_path", "")
        .Range("AV5:AV18").WrapText = True
    End With
End Sub

Private Sub AppendStatusLog(ByVal info As Object)
    Dim stamp As String
    stamp = NzDict(info, "updated", "")
    If stamp = "" Then Exit Sub
    If stamp = lastLoggedStamp Then Exit Sub
    lastLoggedStamp = stamp

    Dim ws As Worksheet
    Set ws = GetSheet(DASHBOARD_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim tbl As ListObject
    Set tbl = EnsureStatusTable(ws)
    If tbl Is Nothing Then Exit Sub

    Dim row As ListRow
    Set row = tbl.ListRows.Add
    With row.Range
        .Cells(1, 1).Value = NzDict(info, "updated", "")
        .Cells(1, 2).Value = NzDict(info, "state", "")
        .Cells(1, 3).Value = NzDict(info, "step", "")
        .Cells(1, 4).Value = NzDict(info, "message", "")
    End With

    If tbl.ListRows.Count > 200 Then
        tbl.ListRows(1).Delete
    End If
End Sub

Private Function EnsureStatusTable(ByVal ws As Worksheet) As ListObject
    On Error Resume Next
    Set EnsureStatusTable = ws.ListObjects(LOG_TABLE_NAME)
    On Error GoTo 0
    If EnsureStatusTable Is Nothing Then
        Dim headerRange As Range
        Set headerRange = ws.Range("AU20:AX20")
        headerRange.Value = Array("Timestamp", "State", "Step", "Message")
        Set EnsureStatusTable = ws.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
        EnsureStatusTable.Name = LOG_TABLE_NAME
        On Error Resume Next
        EnsureStatusTable.TableStyle = "TableStyleLight9"
        On Error GoTo 0
    End If
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function NzDict(ByVal dict As Object, ByVal key As String, ByVal fallback As String) As String
    On Error GoTo FailLookup
    If dict Is Nothing Then GoTo FailLookup
    NzDict = CStr(dict(key))
    Exit Function

FailLookup:
    NzDict = fallback
    Err.Clear
End Function

Private Function FormatElapsedSeconds(ByVal secondsValue As String) As String
    If Len(secondsValue) = 0 Then Exit Function
    If Not IsNumeric(secondsValue) Then
        FormatElapsedSeconds = secondsValue
        Exit Function
    End If

    Dim totalSeconds As Long
    totalSeconds = CLng(secondsValue)
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long

    hours = totalSeconds \ 3600
    minutes = (totalSeconds Mod 3600) \ 60
    seconds = totalSeconds Mod 60

    If hours > 0 Then
        FormatElapsedSeconds = hours & "h " & Format$(minutes, "00") & "m " & Format$(seconds, "00") & "s"
    ElseIf minutes > 0 Then
        FormatElapsedSeconds = minutes & "m " & Format$(seconds, "00") & "s"
    Else
        FormatElapsedSeconds = seconds & "s"
    End If
End Function
