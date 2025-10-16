Attribute VB_Name = "CheckEnv"
Sub CheckEnv()
    Dim p As String: p = ThisWorkbook.Worksheets("Settings").Range("B1").Value
    Dim lastD As Long: lastD = Worksheets("Dashboard").Cells(Rows.Count, "A").End(xlUp).Row
    MsgBox "Settings!B1=" & IIf(Len(p) = 0, "<‹ó>", p) & vbCrLf & _
           "Dashboard rows=" & lastD & vbCrLf & _
           "Calc=" & Application.Calculation, vbInformation
End Sub


