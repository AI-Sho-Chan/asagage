Option Explicit

' Morning automation for ASAGAKE.xlsm
' - Avoid leaving hidden Excel instances that lock the workbook
' - Logs to C:\AI\asagake\logs\morning_task_*.log
'
' NOTE:
' This task must run in the interactive user session.
' If it runs "whether user is logged on or not", Excel may start in Session 0
' (no window) and can lock ASAGAKE.xlsm, causing read-only opens.
'
' IMPORTANT:
' This script intentionally leaves Excel open (so you can see it),
' but ensures this .vbs itself exits quickly and does not hang forever.

Const BASE_DIR = "C:\AI\asagake"
Const LOG_DIR = "C:\AI\asagake\logs"
Const WORKBOOK_PATH = "C:\AI\asagake\ASAGAKE.xlsm"
Const LOCK_PATH = "C:\AI\asagake\logs\morning_task.lock"
Const MIN_CANDIDATE_ROWS = 10

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(LOG_DIR) Then
    On Error Resume Next
    fso.CreateFolder LOG_DIR
    On Error GoTo 0
End If

Dim stamp: stamp = NowStamp()
Dim logPath: logPath = LOG_DIR & "\morning_task_" & stamp & ".log"
Dim logLatest: logLatest = LOG_DIR & "\morning_task_latest.log"

If Not AcquireLock(LOCK_PATH, 60) Then
    LogLine logPath, logLatest, "skip: lock exists (recent)."
    WScript.Quit 0
End If

On Error Resume Next

LogLine logPath, logLatest, "start: workbook=" & WORKBOOK_PATH

' Safety guard:
' This script MUST NOT run under SYSTEM/NT AUTHORITY.
' If it does, Excel may launch in a non-interactive session and lock ASAGAKE.xlsm without showing a window.
Dim guardShell: Set guardShell = CreateObject("WScript.Shell")
Dim envUser: envUser = UCase$(Trim$(guardShell.ExpandEnvironmentStrings("%USERNAME%")))
Dim envDomain: envDomain = UCase$(Trim$(guardShell.ExpandEnvironmentStrings("%USERDOMAIN%")))
Dim envSession: envSession = UCase$(Trim$(guardShell.ExpandEnvironmentStrings("%SESSIONNAME%")))
If envDomain = "NT AUTHORITY" Or envUser = "SYSTEM" Or envSession = "SERVICES" Then
    LogLine logPath, logLatest, "fatal: refuse to run in non-interactive context (USERDOMAIN=" & envDomain & " USERNAME=" & envUser & " SESSIONNAME=" & envSession & ")"
    LogLine logPath, logLatest, "action: fix Task Scheduler to run as logged-in user (InteractiveToken) only."
    ReleaseLock LOCK_PATH
    WScript.Quit 4
End If

If Not fso.FileExists(WORKBOOK_PATH) Then
    LogLine logPath, logLatest, "error: workbook not found"
    ReleaseLock LOCK_PATH
    WScript.Quit 2
End If

' Log candidate CSV size (helps diagnose "only 1 ticker imported" cases)
LogCandidateSummary logPath, logLatest, BASE_DIR & "\output\excel\candidates_nextday.csv"
If Err.Number <> 0 Then
    LogLine logPath, logLatest, "warn: LogCandidateSummary failed: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
End If
Dim candApprox: candApprox = CandidateApproxRecords(BASE_DIR & "\output\excel\candidates_nextday.csv")
Dim candPath: candPath = BASE_DIR & "\output\excel\candidates_nextday.csv"
Dim candLastGoodPath: candLastGoodPath = BASE_DIR & "\output\excel\candidates_nextday_last_good.csv"
If candApprox < MIN_CANDIDATE_ROWS Then
    LogLine logPath, logLatest, "warn: candidates_nextday too small (records~=" & CStr(candApprox) & "); attempt restore from last_good"
    If RestoreFileIfExists(candLastGoodPath, candPath, logPath, logLatest) Then
        candApprox = CandidateApproxRecords(candPath)
        LogLine logPath, logLatest, "info: restored candidates_nextday from last_good; records~=" & CStr(candApprox)
    Else
        LogLine logPath, logLatest, "warn: restore candidates_nextday failed; skip ImportCandidatesV2 and StartDemoV2"
    End If
End If

Dim shell: Set shell = CreateObject("WScript.Shell")
Dim excelApp: Set excelApp = Nothing
Dim createdExcel: createdExcel = False

' Always create a dedicated Excel instance for ASAGAKE.
' Rationale:
' - If we attach to an existing Excel instance, ASAGAKE's AutoTickV2 (timer VBA)
'   will periodically block *all* workbooks inside that same Excel process.
' - A dedicated instance isolates ASAGAKE so other Excel work is not "busy".
Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    LogLine logPath, logLatest, "fatal: CreateObject(Excel.Application) failed: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
    ReleaseLock LOCK_PATH
    WScript.Quit 1
End If
createdExcel = True

On Error Resume Next
' Isolation (important):
' Prevent other Excel files (double-click etc.) from being opened into *this* Excel instance.
' This avoids "other workbooks become busy" when ASAGAKE is auto-ticking every few seconds.
excelApp.IgnoreRemoteRequests = True
On Error GoTo 0

excelApp.Visible = True
excelApp.DisplayAlerts = False
excelApp.AskToUpdateLinks = False
excelApp.UserControl = True
On Error Resume Next
excelApp.WindowState = -4143 ' xlNormal
On Error GoTo 0

Dim wb: Set wb = FindOpenWorkbook(excelApp, WORKBOOK_PATH)
If Err.Number <> 0 Then
    LogLine logPath, logLatest, "fatal: FindOpenWorkbook failed: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
    If createdExcel Then excelApp.Quit
    ReleaseLock LOCK_PATH
    WScript.Quit 1
End If
If wb Is Nothing Then
    Set wb = excelApp.Workbooks.Open(WORKBOOK_PATH, 0, False)
    If Err.Number <> 0 Then
        LogLine logPath, logLatest, "fatal: Workbooks.Open failed: " & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        If createdExcel Then excelApp.Quit
        ReleaseLock LOCK_PATH
        WScript.Quit 1
    End If
End If

wb.Activate
shell.AppActivate wb.Name
If Err.Number <> 0 Then
    LogLine logPath, logLatest, "warn: AppActivate failed: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
End If

On Error Resume Next
excelApp.WindowState = -4137 ' xlMaximized
excelApp.ActiveWindow.WindowState = -4137
On Error GoTo 0

If wb.ReadOnly Then
    LogLine logPath, logLatest, "error: workbook opened as ReadOnly (likely locked by another Excel)."
    LogLine logPath, logLatest, "action: close ReadOnly copy and exit."
    wb.Close False
    If createdExcel Then excelApp.Quit
    ReleaseLock LOCK_PATH
    WScript.Quit 3
End If

If candApprox >= MIN_CANDIDATE_ROWS Then
    LogLine logPath, logLatest, "run: ImportCandidatesV2"
    excelApp.Run "'" & wb.Name & "'!AutoTraderAdvanced.ImportCandidatesV2"
    If Err.Number <> 0 Then
        LogLine logPath, logLatest, "fatal: ImportCandidatesV2 failed: " & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        ReleaseLock LOCK_PATH
        WScript.Quit 1
    End If
Else
    LogLine logPath, logLatest, "skip: ImportCandidatesV2 (candidates too small)"
End If

' Wait a bit for formulas/RSS to settle (keep short to avoid task hanging)
WScript.Sleep 30000

LogLine logPath, logLatest, "run: StartDemoV2"
If candApprox >= MIN_CANDIDATE_ROWS Then
    excelApp.Run "'" & wb.Name & "'!AutoTraderAdvanced.StartDemoV2"
    If Err.Number <> 0 Then
        LogLine logPath, logLatest, "fatal: StartDemoV2 failed: " & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        ReleaseLock LOCK_PATH
        WScript.Quit 1
    End If
Else
    LogLine logPath, logLatest, "skip: StartDemoV2 (candidates too small)"
End If

LogLine logPath, logLatest, "done: leaving Excel open"
ReleaseLock LOCK_PATH
WScript.Quit 0

' ===== helpers =====

Function RestoreFileIfExists(ByVal srcPath, ByVal dstPath, ByVal logPath, ByVal logLatest)
    RestoreFileIfExists = False
    On Error Resume Next
    If Not fso.FileExists(srcPath) Then Exit Function

    ' Create destination folder if needed (should already exist).
    Dim parentDir
    parentDir = fso.GetParentFolderName(dstPath)
    If Len(parentDir) > 0 And Not fso.FolderExists(parentDir) Then
        fso.CreateFolder parentDir
    End If

    ' Copy overwrite.
    fso.CopyFile srcPath, dstPath, True
    If Err.Number <> 0 Then
        LogLine logPath, logLatest, "warn: RestoreFileIfExists copy failed: " & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        Exit Function
    End If

    RestoreFileIfExists = True
End Function

Function NowStamp()
    Dim d: d = Now
    NowStamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & _
               Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Sub LogLine(ByVal pathA, ByVal pathB, ByVal msg)
    On Error Resume Next
    Dim ts: ts = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                 Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
    Dim line: line = ts & " " & msg
    AppendFile pathA, line
    AppendFile pathB, line
End Sub

Sub AppendFile(ByVal path, ByVal line)
    ' Always write *something* to disk (avoid silent failure in scheduled tasks).
    ' Use ANSI (TristateFalse=0) to keep it simple and robust.
    On Error Resume Next

    Dim localFso: Set localFso = CreateObject("Scripting.FileSystemObject")
    Dim f: Set f = Nothing

    Err.Clear
    Set f = localFso.OpenTextFile(path, 8, True, 0) ' ForAppending, create, ANSI
    If Err.Number <> 0 Then
        Err.Clear
    Set f = localFso.CreateTextFile(path, False, False) ' overwrite=False, Unicode=False
    End If

    If Not f Is Nothing Then
        f.WriteLine line
        f.Close
    End If
End Sub

Function AcquireLock(ByVal lockPath, ByVal maxAgeMinutes)
    AcquireLock = False
    On Error Resume Next
    If fso.FileExists(lockPath) Then
        Dim ageMin
        ageMin = DateDiff("n", fso.GetFile(lockPath).DateLastModified, Now)
        If ageMin < maxAgeMinutes Then
            Exit Function
        End If
        fso.DeleteFile lockPath, True
    End If
    Dim f: Set f = fso.OpenTextFile(lockPath, 2, True) ' ForWriting
    f.WriteLine NowStamp()
    f.Close
    AcquireLock = True
End Function

Sub ReleaseLock(ByVal lockPath)
    On Error Resume Next
    If fso.FileExists(lockPath) Then fso.DeleteFile lockPath, True
End Sub

Function FindOpenWorkbook(ByVal app, ByVal fullPath)
    On Error Resume Next
    Dim w
    For Each w In app.Workbooks
        If LCase(w.FullName) = LCase(fullPath) Then
            Set FindOpenWorkbook = w
            Exit Function
        End If
    Next
    Set FindOpenWorkbook = Nothing
End Function

Sub LogCandidateSummary(ByVal pathA, ByVal pathB, ByVal csvPath)
    On Error Resume Next
    If Not fso.FileExists(csvPath) Then
        LogLine pathA, pathB, "candidates_nextday: missing (" & csvPath & ")"
        Exit Sub
    End If

    Dim raw: raw = ReadAllTextUtf8BestEffort(csvPath)

    Dim sizeBytes: sizeBytes = 0
    Dim mtime: mtime = ""
    On Error Resume Next
    sizeBytes = fso.GetFile(csvPath).Size
    mtime = CStr(fso.GetFile(csvPath).DateLastModified)
    On Error GoTo 0

    If Len(raw) = 0 And sizeBytes > 0 Then
        LogLine pathA, pathB, "warn: candidates_nextday read returned empty (transient lock?) size=" & CStr(sizeBytes) & " mtime=" & mtime
    End If

    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    Dim lines: lines = Split(raw, vbLf)
    Dim n: n = UBound(lines) + 1
    Dim records: records = n - 1
    If records < 0 Then records = 0

    LogLine pathA, pathB, "candidates_nextday: lines=" & CStr(n) & " records~=" & CStr(records) & " size=" & CStr(sizeBytes) & " mtime=" & mtime & " (" & csvPath & ")"
End Sub

Function CandidateApproxRecords(ByVal csvPath)
    On Error Resume Next
    CandidateApproxRecords = 0
    If Not fso.FileExists(csvPath) Then Exit Function

    Dim raw: raw = ReadAllTextUtf8BestEffort(csvPath)
    If Len(raw) = 0 Then
        ' If the read fails transiently but the file is clearly non-empty, do not treat it as 0 records.
        Dim sizeBytes: sizeBytes = 0
        On Error Resume Next
        sizeBytes = fso.GetFile(csvPath).Size
        On Error GoTo 0
        If sizeBytes > 1024 Then
            CandidateApproxRecords = MIN_CANDIDATE_ROWS
            Exit Function
        End If
    End If
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    Dim lines: lines = Split(raw, vbLf)
    Dim n: n = UBound(lines) + 1
    Dim records: records = n - 1
    If records < 0 Then records = 0
    CandidateApproxRecords = records
End Function

Function ReadAllTextUtf8BestEffort(ByVal path)
    On Error Resume Next

    ' Prefer ADODB.Stream (proper UTF-8 + BOM handling).
    Dim stm: Set stm = CreateObject("ADODB.Stream")
    If Err.Number = 0 And Not stm Is Nothing Then
        Err.Clear
        stm.Type = 2 ' text
        stm.Charset = "utf-8"
        stm.Open
        Dim i
        For i = 1 To 3
            Err.Clear
            stm.LoadFromFile path
            If Err.Number = 0 Then Exit For
            WScript.Sleep 150
        Next
        If Err.Number = 0 Then
            ReadAllTextUtf8BestEffort = stm.ReadText(-1)
            stm.Close
            Exit Function
        End If
        ' If LoadFromFile failed (e.g. transient lock), fall back instead of returning empty.
        Err.Clear
        stm.Close
    End If
    Err.Clear

    ' Fallback: FSO read (may mis-detect encoding; still better than failing).
    Dim f: Set f = fso.OpenTextFile(path, 1, False, 0) ' ForReading, ANSI
    If Err.Number <> 0 Then
        Err.Clear
        ReadAllTextUtf8BestEffort = ""
        Exit Function
    End If

    If Not f.AtEndOfStream Then
        ReadAllTextUtf8BestEffort = f.ReadAll
    Else
        ReadAllTextUtf8BestEffort = ""
    End If
    f.Close
End Function
