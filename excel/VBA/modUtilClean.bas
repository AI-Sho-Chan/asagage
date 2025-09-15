Attribute VB_Name = "modUtilClean"
'--- modUtilClean.bas ---
Option Explicit
Public Function CleanTicker(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    ' ���䕶���̏���
    s = WorksheetFunction.Clean(s)
    ' �m�[�u���[�N�X�y�[�X / �S�p�X�y�[�X / �^�u / ���s������
    s = Replace(s, ChrW(160), "")
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim$(s)
    ' ���l�Ȃ� "0" �`���̕������
    If Len(s) > 0 Then
        If IsNumeric(s) Then s = Format$(CLng(s), "0")
    End If
    CleanTicker = s
End Function

