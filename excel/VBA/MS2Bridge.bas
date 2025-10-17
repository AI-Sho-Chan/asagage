Attribute VB_Name = "MS2Bridge"
Option Explicit

' Bridge macro to call Rakuten RSS order functions from NewDashboard
' This is a safe placeholder: logs into Orders and is ready to call RSS once mapping is finalized.

Public Sub Place(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal info As String)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("NewDashboard")
    Dim qty As Long
    qty = CLng(Nz(wsDash.Range("B12").Value, 0))
    If qty <= 0 Then qty = 100

    Dim tif As String
    tif = CStr(Nz(wsDash.Range("B13").Value, "MKT"))

    Dim kind As String
    kind = UCase$(info)
    If kind = "ENTRY" Then
        Call EvalTemplate("EntryTemplate", ticker, side, qty, price, tif)
    ElseIf kind = "TP" Then
        Call EvalTemplate("TPTemplate", ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), qty, price, tif)
    ElseIf kind = "SL" Then
        Call EvalTemplate("SLTemplate", ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), qty, price, tif)
    ElseIf kind = "MOC" Or kind = "FLAT" Then
        Call EvalTemplate("MOCTemplate", ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), qty, price, tif)
    Else
        LogOrder ticker, side, qty, price, info
    End If
End Sub

Private Sub EvalTemplate(ByVal key As String, ByVal ticker As String, ByVal side As String, ByVal qty As Long, ByVal price As Variant, ByVal tif As String)
    Dim tmpl As String
    tmpl = GetConfig(key)
    If Len(tmpl) = 0 Then
        LogOrder ticker, side, qty, price, key & ":NO_TEMPLATE"
        Exit Sub
    End If
    Dim expr As String
    expr = tmpl
    expr = Replace(expr, "{Ticker}", ticker)
    expr = Replace(expr, "{Side}", side)
    expr = Replace(expr, "{Qty}", CStr(qty))
    expr = Replace(expr, "{Price}", IIf(IsNumeric(price), CStr(price), ""))
    expr = Replace(expr, "{TIF}", tif)
    expr = Replace(expr, "{Account}", GetConfig("Account"))
    expr = Replace(expr, "{Market}", GetConfig("Market"))
    On Error Resume Next
    Dim result
    result = Application.Evaluate("=" & expr)
    If Err.Number <> 0 Then
        LogOrder ticker, side, qty, price, key & ":EVAL_ERR:" & CStr(Err.Description)
        Err.Clear
    Else
        LogOrder ticker, side, qty, price, key & ":EVAL_OK"
    End If
End Sub

Private Function GetConfig(ByVal key As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet("MS2_Config")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To lastRow
        If CStr(ws.Cells(r, 1).Value) = key Then
            GetConfig = CStr(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    GetConfig = ""
End Function

Private Sub LogOrder(ByVal ticker As String, ByVal side As String, ByVal qty As Long, ByVal price As Variant, ByVal info As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet("Orders")
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r = 1 Then
        ws.Range("A1:F1").Value = Array("Time", "Ticker", "Side", "Qty", "Price", "Note")
        r = 2
    End If
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = ticker
    ws.Cells(r, 3).Value = side
    ws.Cells(r, 4).Value = qty
    ws.Cells(r, 5).Value = price
    ws.Cells(r, 6).Value = "MS2Bridge:" & info
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
