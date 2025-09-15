Attribute VB_Name = "modUtilClean"
'--- modUtilClean.bas ---
Option Explicit
Public Function CleanTicker(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    ' 制御文字の除去
    s = WorksheetFunction.Clean(s)
    ' ノーブレークスペース / 全角スペース / タブ / 改行を除去
    s = Replace(s, ChrW(160), "")
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim$(s)
    ' 数値なら "0" 形式の文字列に
    If Len(s) > 0 Then
        If IsNumeric(s) Then s = Format$(CLng(s), "0")
    End If
    CleanTicker = s
End Function

