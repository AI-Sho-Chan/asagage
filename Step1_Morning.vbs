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
Const MIN_CANDIDATE_ROWS = 5

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
If candApprox < MIN_CANDIDATE_ROWS Then
    LogLine logPath, logLatest, "warn: candidates_nextday too small (records~=" & CStr(candApprox) & "); skip ImportCandidatesV2 and StartDemoV2"
End If

Dim shell: Set shell = CreateObject("WScript.Shell")
Dim excelApp: Set excelApp = Nothing
Dim createdExcel: createdExcel = False

' Try attach (same desktop/session) first
Set excelApp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Err.Clear
End If
If excelApp Is Nothing Then
    Set excelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        LogLine logPath, logLatest, "fatal: CreateObject(Excel.Application) failed: " & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        ReleaseLock LOCK_PATH
        WScript.Quit 1
    End If
    createdExcel = True
End If

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

    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    Dim lines: lines = Split(raw, vbLf)
    Dim n: n = UBound(lines) + 1
    Dim records: records = n - 1
    If records < 0 Then records = 0

    Dim sizeBytes: sizeBytes = 0
    Dim mtime: mtime = ""
    On Error Resume Next
    sizeBytes = fso.GetFile(csvPath).Size
    mtime = CStr(fso.GetFile(csvPath).DateLastModified)
    On Error GoTo 0

    LogLine pathA, pathB, "candidates_nextday: lines=" & CStr(n) & " records~=" & CStr(records) & " size=" & CStr(sizeBytes) & " mtime=" & mtime & " (" & csvPath & ")"
End Sub

Function CandidateApproxRecords(ByVal csvPath)
    On Error Resume Next
    CandidateApproxRecords = 0
    If Not fso.FileExists(csvPath) Then Exit Function

    Dim raw: raw = ReadAllTextUtf8BestEffort(csvPath)
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
        stm.Type = 2 ' text
        stm.Charset = "utf-8"
        stm.Open
        stm.LoadFromFile path
        ReadAllTextUtf8BestEffort = stm.ReadText(-1)
        stm.Close
        Exit Function
    End If
    Err.Clear

    ' Fallback: FSO read (may mis-detect encoding; still better than failing).
    Dim f: Set f = fso.OpenTextFile(path, 1, False, 0) ' ForReading, ANSI
    If Not f.AtEndOfStream Then
        ReadAllTextUtf8BestEffort = f.ReadAll
    Else
        ReadAllTextUtf8BestEffort = ""
    End If
    f.Close
End Function
