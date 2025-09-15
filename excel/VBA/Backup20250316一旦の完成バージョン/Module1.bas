Attribute VB_Name = "Module1"
Option Explicit

Public Function Nz(x As Variant, Optional d As Double = 0#) As Double
    If IsError(x) Then
        Nz = d
    ElseIf IsNumeric(x) Then
        Nz = CDbl(x)
    Else
        Nz = d
    End If
End Function

Public Function CleanTicker(v As Variant) As String
    Dim s As String, i As Long, t As String, ch As String
    s = CStr(v)
    s = Replace(s, "Å@", " ")
    s = Replace(s, ChrW(12288), " ")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then t = t & ch
    Next
    If Len(t) = 0 Then t = s
    CleanTicker = t
End Function

Public Sub EnsureDir(p As String)
    If Len(dir(p, vbDirectory)) = 0 Then MkDir p
End Sub

Public Sub AppendCsv(ByVal fileName As String, ByVal line As String)
    Dim ff As Integer: ff = FreeFile
    Open fileName For Append As #ff
    Print #ff, line
    Close #ff
End Sub

