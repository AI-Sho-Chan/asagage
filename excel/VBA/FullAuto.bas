Attribute VB_Name = "FullAuto"
Option Explicit

Public Sub Place(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal info As String)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("NewDashboard")
    Dim qty As Long
    qty = CLng(Nz(wsDash.Range("B12").Value, 0))
    If qty <= 0 Then qty = 100

    Dim tif As String
    tif = CStr(Nz(wsDash.Range("B13").Value, "MKT"))

    Dim ordType As String
    If UCase$(info) = "MOC" Then
        ordType = "MOC"
    ElseIf IsNumeric(price) And CDbl(price) > 0 Then
        ordType = "LMT"
    Else
        ordType = "MKT"
    End If

    Dim wsLog As Worksheet
    Set wsLog = EnsureSheet("Orders")
    Dim r As Long
    r = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If r = 1 Then
        wsLog.Range("A1:F1").Value = Array("Time", "Ticker", "Side", "Qty", "Price", "Note")
        r = 2
    End If
    wsLog.Cells(r, 1).Value = Now
    wsLog.Cells(r, 2).Value = ticker
    wsLog.Cells(r, 3).Value = side
    wsLog.Cells(r, 4).Value = qty
    wsLog.Cells(r, 5).Value = price
    wsLog.Cells(r, 6).Value = ordType & ":" & info

    ' Hook to Rakuten RSS order macro/function here if available
    ' Example: Application.Run "RSS_Order.Place", ticker, side, qty, price, ordType, tif, info
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function Nz(ByVal v As Variant, ByVal def As Variant) As Variant
    If IsError(v) Then
        Nz = def
    ElseIf IsEmpty(v) Then
        Nz = def
    ElseIf v = "" Then
        Nz = def
    Else
        Nz = v
    End If
End Function

