Attribute VB_Name = "M_ASG_Core"
Option Explicit

'=== constants ===
Private Const R2   As Long = 2      'Bars header row
Private Const R3   As Long = 3      'data start row
Private Const MAXN As Long = 300    'max blocks
Private Const BLKW As Long = 12     'block width B..K
Private nextTick   As Date

'=== helpers ===
Private Function BlockCol(ByVal idx As Long) As Long
    BlockCol = 2 + (idx - 1) * BLKW   'B=2, N=14, Z=26, ...
End Function

Private Function Code4(ByVal v As Variant) As String
    Dim s As String, t As String, i As Long, ch As String
    s = CStr(v)
    s = Replace(Replace(Replace(Replace(Replace(s, "@", ""), " ", ""), "-", ""), "/", ""), "_", "")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then t = t & ch
    Next
    If Len(t) = 4 Then Code4 = t Else Code4 = ""
End Function

Public Function GetUniverse(ByRef codes() As String) As Long
    Dim wsD As Worksheet: Set wsD = Worksheets("Dashboard")
    Dim last As Long, r As Long, v As Variant, n As Long
    last = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    ReDim codes(1 To MAXN): n = 0
    For r = 2 To last
        v = Code4(wsD.Cells(r, "A").Value2)
        If Len(v) = 4 Then n = n + 1: codes(n) = v: If n = MAXN Then Exit For
    Next
    If n > 0 Then ReDim Preserve codes(1 To n)
    GetUniverse = n
End Function

'=== Bars block locator ===
Public Function FindBlockCol(ByVal code As String) As Long
    Dim wsB As Worksheet: Set wsB = Worksheets("Bars")
    Dim i As Long, sc As Long, fx As String
    For i = 1 To MAXN
        sc = BlockCol(i)
        fx = CStr(wsB.Cells(R2, sc - 1).Formula2)
        If Len(fx) = 0 Then Exit For
        fx = Replace(fx, "=@RssChart", "=RssChart")
        If InStr(1, fx, """" & code & """", vbTextCompare) > 0 Then FindBlockCol = sc: Exit Function
    Next
End Function

Public Function LastRowInBlock(ByVal sc As Long) As Long
    Dim ws As Worksheet: Set ws = Worksheets("Bars")
    Dim j As Long, lastCol As Long, mx As Long
    mx = R3
    For j = sc To sc + 9
        lastCol = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
        If lastCol > mx Then mx = lastCol
    Next
    LastRowInBlock = mx
End Function


