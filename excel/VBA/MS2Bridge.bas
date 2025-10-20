Attribute VB_Name = "MS2Bridge"
Option Explicit

' Bridge macro to call Rakuten RSS order functions from NewDashboard
' This is a safe placeholder: logs into Orders and is ready to call RSS once mapping is finalized.

Public Sub Place(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("NewDashboard")
    If qty <= 0 Then
        qty = CLng(Nz(wsDash.Range("B12").Value, 100))
    End If

    Dim tif As String
    tif = CStr(Nz(wsDash.Range("B13").Value, "MKT"))

    Dim kind As String
    kind = UCase$(CStr(info))
    Dim targetSide As String
    Select Case kind
        Case "ENTRY"
            targetSide = side
            ExecuteTemplate "EntryTemplate", ticker, targetSide, qty, price, tif, kind
        Case "TP"
            targetSide = IIf(UCase$(side) = "BUY", "SELL", "BUY")
            ExecuteTemplate "TPTemplate", ticker, targetSide, qty, price, tif, kind
        Case "SL"
            targetSide = IIf(UCase$(side) = "BUY", "SELL", "BUY")
            ExecuteTemplate "SLTemplate", ticker, targetSide, qty, price, tif, kind
        Case "MOC", "FLAT"
            targetSide = IIf(UCase$(side) = "BUY", "SELL", "BUY")
            ExecuteTemplate "MOCTemplate", ticker, targetSide, qty, price, tif, kind
        Case Else
            LogOrder ticker, side, qty, price, "UNKNOWN:" & kind
    End Select
End Sub

Private Sub ExecuteTemplate(ByVal key As String, ByVal ticker As String, ByVal side As String, ByVal qty As Long, ByVal price As Variant, ByVal tif As String, ByVal info As String)
    Dim cfg As Object
    Set cfg = LoadTemplateRow(key)
    If cfg Is Nothing Then
        LogOrder ticker, side, qty, price, key & ":NO_TEMPLATE"
        Exit Sub
    End If

    Dim expr As String
    expr = BuildExpression(cfg, ticker, side, qty, price, tif, info)
    If Len(expr) = 0 Then
        LogOrder ticker, side, qty, price, key & ":INVALID_TEMPLATE"
        Exit Sub
    End If

    On Error GoTo EvalFail
    Dim result
    result = Application.Evaluate("=" & expr)
    On Error GoTo 0
    LogOrder ticker, side, qty, price, key & ":EVAL_OK"
    Exit Sub

EvalFail:
    LogOrder ticker, side, qty, price, key & ":EVAL_ERR:" & Err.Description
    Err.Clear
End Sub

Private Function LoadTemplateRow(ByVal key As String) As Object
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet("MS2_Config")
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 2 Then Exit Function

    Dim headers As Variant
    headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value

    Dim r As Long, c As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 1).Value) = key Then
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            For c = 1 To lastCol
                Dim headerName As String
                headerName = Trim$(CStr(headers(1, c)))
                If Len(headerName) > 0 Then
                    dict(headerName) = ws.Cells(r, c).Value
                End If
            Next c
            Set LoadTemplateRow = dict
            Exit Function
        End If
    Next r
End Function

Private Function BuildExpression(ByVal cfg As Object, ByVal ticker As String, ByVal side As String, ByVal qty As Long, ByVal price As Variant, ByVal tif As String, ByVal info As String) As String
    If cfg Is Nothing Then
        BuildExpression = ""
        Exit Function
    End If

    Dim templateValue As String
    templateValue = CStr(NzDict(cfg, "Value", ""))
    If Len(templateValue) = 0 Then
        BuildExpression = ""
        Exit Function
    End If

    Dim replacements As Object
    Set replacements = CreateObject("Scripting.Dictionary")

    Dim orderIdFmt As String
    orderIdFmt = CStr(NzDict(cfg, "OrderIdFmt", "{Ticker}-{Info}-{Time}"))
    replacements("OrderId") = QuoteLiteral(ExpandOrderId(orderIdFmt, ticker, side, qty, price, info))

    replacements("Ticker") = QuoteLiteral(ticker)
    replacements("TickerCode") = QuoteLiteral(ticker)
    replacements("Side") = QuoteLiteral(UCase$(side))
    replacements("Qty") = CStr(qty)
    replacements("Info") = QuoteLiteral(info)
    replacements("TIF") = QuoteLiteral(tif)
    replacements("Account") = QuoteLiteral(GetConfig("Account"))
    replacements("Market") = QuoteLiteral(GetConfig("Market"))
    replacements("Date") = QuoteLiteral(Format(Date, "yyyymmdd"))
    replacements("Time") = QuoteLiteral(Format(Time, "HHMMSS"))

    Dim buyCode As String
    Dim sellCode As String
    buyCode = FormatNumericValue(NzDict(cfg, "BuyCode", 1), 1)
    sellCode = FormatNumericValue(NzDict(cfg, "SellCode", 2), 2)
    If UCase$(side) = "BUY" Then
        replacements("SideCode") = buyCode
    Else
        replacements("SideCode") = sellCode
    End If

    replacements("OrderDiv") = FormatNumericValue(NzDict(cfg, "OrderDiv", 1), 1)
    replacements("SorDiv") = FormatNumericValue(NzDict(cfg, "SorDiv", 0), 0)
    replacements("CreditDiv") = FormatOptionalNumeric(NzDict(cfg, "CreditDiv", ""))
    replacements("PriceDiv") = FormatNumericValue(NzDict(cfg, "PriceDiv", 0), 0)
    replacements("ExecCond") = FormatNumericValue(NzDict(cfg, "ExecCond", 0), 0)
    replacements("Term") = FormatNumericValue(NzDict(cfg, "Term", 0), 0)
    replacements("AccountDiv") = FormatNumericValue(NzDict(cfg, "AccountDiv", 0), 0)

    Dim orderPrice As String
    If IsNumeric(price) And CDbl(price) > 0 Then
        orderPrice = CStr(price)
    Else
        orderPrice = FormatOptionalNumeric(NzDict(cfg, "DefaultPrice", ""))
    End If
    replacements("OrderPrice") = orderPrice

    replacements("TriggerPrice1") = FormatOptionalNumeric(NzDict(cfg, "TriggerPrice1", ""))
    replacements("TriggerCond1") = FormatNumericValue(NzDict(cfg, "TriggerCond1", 0), 0)
    replacements("TriggerPrice2") = FormatOptionalNumeric(NzDict(cfg, "TriggerPrice2", ""))
    replacements("TriggerCond2") = FormatNumericValue(NzDict(cfg, "TriggerCond2", 0), 0)
    replacements("SetDiv") = FormatNumericValue(NzDict(cfg, "SetDiv", 0), 0)
    replacements("SetPriceDiv") = FormatNumericValue(NzDict(cfg, "SetPriceDiv", 0), 0)
    replacements("SetPrice") = FormatOptionalNumeric(NzDict(cfg, "SetPrice", ""))
    replacements("SetExecCond") = FormatNumericValue(NzDict(cfg, "SetExecCond", 0), 0)
    replacements("SetAccount") = FormatNumericValue(NzDict(cfg, "SetAccount", 0), 0)

    replacements("Price") = FormatOptionalNumeric(price)

    Dim expr As String
    expr = ApplyReplacements(templateValue, replacements)
    BuildExpression = expr
End Function

Private Function ApplyReplacements(ByVal template As String, ByVal replacements As Object) As String
    Dim result As String
    result = template
    Dim key As Variant
    For Each key In replacements.Keys
        result = Replace(result, "{" & CStr(key) & "}", CStr(replacements(key)))
    Next key
    ApplyReplacements = result
End Function

Private Function ExpandOrderId(ByVal fmt As String, ByVal ticker As String, ByVal side As String, ByVal qty As Long, ByVal price As Variant, ByVal info As String) As String
    Dim result As String
    result = fmt
    result = Replace(result, "{Ticker}", ticker)
    result = Replace(result, "{Side}", side)
    result = Replace(result, "{Qty}", CStr(qty))
    result = Replace(result, "{Info}", info)
    If IsNumeric(price) Then
        result = Replace(result, "{Price}", CStr(price))
    Else
        result = Replace(result, "{Price}", "")
    End If
    result = Replace(result, "{Date}", Format(Date, "yyyymmdd"))
    result = Replace(result, "{Time}", Format(Time, "HHMMSS"))
    ExpandOrderId = result
End Function

Private Function QuoteLiteral(ByVal value As String) As String
    QuoteLiteral = """" & Replace(value, """", """""") & """"
End Function

Private Function NzDict(ByVal dict As Object, ByVal key As String, ByVal fallback As Variant) As Variant
    If Not dict Is Nothing Then
        If dict.Exists(key) Then
            NzDict = dict(key)
            Exit Function
        End If
    End If
    NzDict = fallback
End Function

Private Function FormatNumericValue(ByVal value As Variant, ByVal defaultValue As Double) As String
    If IsNumeric(value) Then
        FormatNumericValue = CStr(value)
    ElseIf TypeName(value) = "String" Then
        Dim s As String
        s = Trim$(CStr(value))
        If Len(s) = 0 Then
            FormatNumericValue = CStr(defaultValue)
        ElseIf IsNumeric(s) Then
            FormatNumericValue = s
        Else
            FormatNumericValue = CStr(defaultValue)
        End If
    Else
        FormatNumericValue = CStr(defaultValue)
    End If
End Function

Private Function FormatOptionalNumeric(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then
        FormatOptionalNumeric = QuoteLiteral("")
    ElseIf IsNumeric(value) Then
        FormatOptionalNumeric = CStr(value)
    Else
        Dim s As String
        s = Trim$(CStr(value))
        If Len(s) = 0 Then
            FormatOptionalNumeric = QuoteLiteral("")
        ElseIf s = """""""" Then
            FormatOptionalNumeric = QuoteLiteral("")
        ElseIf IsNumeric(s) Then
            FormatOptionalNumeric = s
        Else
            FormatOptionalNumeric = QuoteLiteral(s)
        End If
    End If
End Function

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
        ws.Range("A1:F1").Value = Array("Time", "Ticker", "Side", "Price", "Qty", "Note")
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
