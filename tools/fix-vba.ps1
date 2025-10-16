param(
  [string]$Workbook = "C:\AI\asagake\excel\AsagageRSS.xlsm",
  [string]$OutDir   = "C:\AI\asagake\excel\VBA\fixed"
)

New-Item -ItemType Directory -Force $OutDir | Out-Null

# ---------- modBarsSetup.bas（ASCII） ----------
$modBarsSetup = @'
Attribute VB_Name = "modBarsSetup"
Option Explicit

Private Function HeaderArray() As Variant
    HeaderArray = Array("Name","Market","Foot","Date","Time","Open","High","Low","Close","Volume")
End Function

Public Sub RebuildBarsAll()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsB As Worksheet: Set wsB = wb.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = wb.Sheets("Dashboard")

    Dim lastRow As Long, blocks As Long
    lastRow = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    blocks = WorksheetFunction.Min(20, WorksheetFunction.Max(0, lastRow - 1))

    Application.ScreenUpdating = False
    wsB.Range(wsB.Cells(2,1), wsB.Cells(22, 12 * blocks + 1)).Clear

    Dim r As Long, sc As Long, k As Long, h As Variant
    For r = 2 To blocks + 1
        sc = 2 + (r - 2) * 12
        h = HeaderArray()
        For k = 0 To UBound(h)
            wsB.Cells(2, sc + k).Value = h(k)
        Next k
    Next r

    Application.ScreenUpdating = True
    NudgeRssTriggers
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "RebuildBarsAll: " & Err.Description, vbExclamation
End Sub
'@

# ---------- modNudgeRssTriggers.bas（ASCII） ----------
$modNudge = @'
Attribute VB_Name = "modNudgeRssTriggers"
Option Explicit

Private Const TOP_ROW As Long = 2
Private Const COLS_PER_BLOCK As Long = 12
Private Const MAX_BLOCKS As Long = 20

Private Sub StripAtIfNeeded(ByVal tgt As Range)
    On Error Resume Next
    Dim f As String: f = tgt.Formula2
    If Len(f) >= 2 Then
        If Left$(f,2) = "=@"
        Then tgt.Formula2 = "=" & Mid$(f,3)
        End If
    End If
End Sub

Private Function Fx(ByVal headAddr As String, ByVal codeAddr As String, ByVal foot As String) As String
    Fx = "=RssChart(" & headAddr & ",IFERROR(TEXT(" & codeAddr & ",""0"")," & codeAddr & "&"""")" & ",""" & foot & """,20)"
End Function

Public Sub FixCalcAndRefresh()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.ShowFormulas = False
    Application.CalculateFullRebuild
    Application.ScreenUpdating = True
End Sub

Public Sub NudgeRssTriggers()
    On Error GoTo EH
    Dim wsB As Worksheet: Set wsB = ThisWorkbook.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("Dashboard")
    Dim foot As String: foot = ThisWorkbook.Sheets("Settings").Range("B4").Value

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    Dim blocks As Long: blocks = WorksheetFunction.Min(MAX_BLOCKS, WorksheetFunction.Max(0, lastRow - 1))

    Dim r As Long, sc As Long, hdr As Range, codeCell As Range, f As String
    For r = 2 To blocks + 1
        sc = 2 + (r - 2) * COLS_PER_BLOCK
        Set hdr     = wsB.Range(wsB.Cells(TOP_ROW, sc), wsB.Cells(TOP_ROW, sc + 9))
        Set codeCell= wsD.Cells(r, "A")
        f = Fx(hdr.Address(False, False), codeCell.Address(False, False), foot)
        With wsB.Cells(TOP_ROW, sc - 1)
            .NumberFormat = "General"
            .ClearContents
            .Formula2 = f
            StripAtIfNeeded wsB.Cells(TOP_ROW, sc - 1)
        End With
    Next r
    Exit Sub
EH:
    MsgBox "NudgeRssTriggers: " & Err.Description, vbExclamation
End Sub
'@

# 書き出しは必ず ASCII
$barsPath  = Join-Path $OutDir "modBarsSetup.bas"
$nudgePath = Join-Path $OutDir "modNudgeRssTriggers.bas"
Set-Content $barsPath  $modBarsSetup -Encoding Ascii
Set-Content $nudgePath $modNudge     -Encoding Ascii

# Excel COM でインポート（既存は削除）
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try{
    $wb = $xl.Workbooks.Open($Workbook)
    $vb = $wb.VBProject

    foreach($name in "modBarsSetup","modNudgeRssTriggers"){
        $comp = $vb.VBComponents | Where-Object { $_.Name -eq $name }
        if($comp){ $vb.VBComponents.Remove($comp) }
    }

    $vb.VBComponents.Import($barsPath)  | Out-Null
    $vb.VBComponents.Import($nudgePath) | Out-Null

    $wb.Save()
}
finally{
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
"Done: modules replaced with ASCII versions at $OutDir"
