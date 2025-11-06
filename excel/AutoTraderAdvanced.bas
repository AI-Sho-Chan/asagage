Attribute VB_Name = "AutoTraderAdvanced"

Option Explicit



' Dashboard V2 constants (ASCII only)

Private Const DASH2_SHEET As String = "NewDashboardV2"

Private Const DASH2_HEADER_ROW As Long = 5

Private Const DASH2_DATA_START As Long = 6



Private Const DASH2_PREPLACE_FRACTION_CELL As String = "B50"

Private Const DQ As String = """"

Private Const MsoShapeRectangle As Long = 1

Private Const MsoAlignCenter As Long = -4108



Private Function HeaderTickerJP() As String: HeaderTickerJP = "Ticker": End Function

Private Function HeaderJValueJP() As String: HeaderJValueJP = "J": End Function

Private Function HeaderJThJP() As String: HeaderJThJP = "J_th": End Function



' Helpers

Private Function FindColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long

    Dim c As Range

    For Each c In ws.Rows(headerRow).Cells

        If Trim$(CStr(c.Value)) = headerName Then

            FindColumn = c.Column

            Exit Function

        End If

    Next c

    FindColumn = 0

End Function



Private Sub SetColumnFormula(ByVal ws As Worksheet, ByVal col As Long, ByVal fillLast As Long, ByVal formulaR1C1 As String)

    If col <= 0 Then Exit Sub

    Dim rng As Range

    Set rng = ws.Range(ws.Cells(DASH2_DATA_START, col), ws.Cells(fillLast, col))

    Dim firstCell As Range

    Set firstCell = rng.Cells(1, 1)

    On Error Resume Next

    firstCell.formulaR1C1 = formulaR1C1

    If Err.Number <> 0 Then

        Err.Clear

        firstCell.FormulaR1C1Local = formulaR1C1

    End If

    If rng.Rows.Count > 1 Then firstCell.AutoFill Destination:=rng

    On Error GoTo 0

End Sub



Private Function BuildR1C1Ref(ByVal sourceCol As Long, ByVal targetCol As Long) As String
    Dim d As Long: d = sourceCol - targetCol
    If d = 0 Then
        BuildR1C1Ref = "RC"
    Else
        BuildR1C1Ref = "RC[" & CStr(d) & "]"
    End If
End Function

Private Function ParseCsvLine(ByVal line As String) As Variant

    Dim values As Collection

    Set values = New Collection

    Dim current As String

    Dim i As Long

    Dim ch As String

    Dim inQuotes As Boolean

    For i = 1 To Len(line)

        ch = Mid$(line, i, 1)

        If ch = """" Then

            If inQuotes And i < Len(line) And Mid$(line, i + 1, 1) = """" Then

                current = current & """"

                i = i + 1

            Else

                inQuotes = Not inQuotes

            End If

        ElseIf ch = "," And Not inQuotes Then

            values.Add current

            current = vbNullString

        Else

            current = current & ch

        End If

    Next i

    values.Add current

    Dim arr() As String

    ReDim arr(0 To values.Count - 1)

    For i = 1 To values.Count

        arr(i - 1) = values(i)

    Next i

    ParseCsvLine = arr

End Function



' ----------------------------------------------------------------------------

' Setup/Install

' ----------------------------------------------------------------------------

Public Sub SetupNewDashboardV2()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.name = DASH2_SHEET

    End If



    Dim r As Long: r = DASH2_HEADER_ROW

    ws.Cells(r, 1).Value = HeaderTickerJP()

    ws.Cells(r, 2).Value = "Selected"

    ws.Cells(r, 3).Value = HeaderJValueJP()

    ws.Cells(r, 4).Value = HeaderJThJP()

    ws.Cells(r, 5).Value = "EntryBuyPx"

    ws.Cells(r, 6).Value = "EntrySellPx"

    ws.Cells(r, 7).Value = "EntrySide"

    ws.Cells(r, 8).Value = "EntryStatus"

    ws.Cells(r, 9).Value = "TP_price"

    ws.Cells(r, 10).Value = "SL_price"

    ws.Cells(r, 11).Value = "StopTrail"

    ws.Cells(r, 12).Value = "BestBid"

    ws.Cells(r, 13).Value = "BestAsk"

    ws.Cells(r, 14).Value = "PrevClose"

    ws.Cells(r, 15).Value = "VWAP"

    ws.Cells(r, 16).Value = "Gap_bp"

    ws.Cells(r, 17).Value = "CorrNKY"

    ws.Cells(r, 18).Value = "OrderQtyPlan"

End Sub



Public Sub InstallRealtimeFormulasV2()

    ' Intentionally minimal (no Rss formulas here; keep offline rule)

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Dim gapCol As Long: gapCol = FindColumn(ws, DASH2_HEADER_ROW, "Gap_bp")

    Dim prevCol As Long: prevCol = FindColumn(ws, DASH2_HEADER_ROW, "PrevClose")

    Dim vwapCol As Long: vwapCol = FindColumn(ws, DASH2_HEADER_ROW, "VWAP")

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If gapCol > 0 And prevCol > 0 And vwapCol > 0 Then

        Dim f As String

        f = "=IF(OR(RC[" & (vwapCol - gapCol) & "]=" & DQ & DQ & ",RC[" & (prevCol - gapCol) & "]=" & DQ & DQ & ")," & DQ & DQ & "," & _

            "(RC[" & (vwapCol - gapCol) & "]-RC[" & (prevCol - gapCol) & "])/RC[" & (prevCol - gapCol) & "]*10000)"

        SetColumnFormula ws, gapCol, lastRow, f

    End If

End Sub



' ----------------------------------------------------------------------------

' Signals and Orders

' ----------------------------------------------------------------------------
Private Const SLIPPAGE_FILE As String = "output\excel\slippage_overrides.csv"

Private Function LoadSlippageOverrides() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fullPath As String
    fullPath = ThisWorkbook.Path & "\" & SLIPPAGE_FILE
    If Not fso.FileExists(fullPath) Then
        Set LoadSlippageOverrides = dict
        Exit Function
    End If

    Const ForReading = 1
    Const TristateTrue = -1

    Dim stream As Object
    On Error Resume Next
    Set stream = fso.OpenTextFile(fullPath, ForReading, False, TristateTrue)
    If Err.Number <> 0 Then
        Err.Clear
        Set LoadSlippageOverrides = dict
        Exit Function
    End If
    On Error GoTo 0

    If Not stream.AtEndOfStream Then
        stream.ReadLine ' skip header
    End If

    Do While Not stream.AtEndOfStream
        Dim line As String
        line = Trim$(stream.ReadLine)
        If Len(line) = 0 Then GoTo ContinueLoop

        Dim parts() As String
        parts = Split(line, ",")
        If UBound(parts) < 3 Then GoTo ContinueLoop

        Dim plan As String: plan = Trim$(parts(0))
        Dim sessionVal As String: sessionVal = Trim$(parts(1))
        Dim bufStr As String: bufStr = Trim$(parts(2))
        Dim bufferBp As Double
        On Error Resume Next
        bufferBp = CDbl(bufStr)
        If Err.Number <> 0 Then
            Err.Clear
            bufferBp = 0
        End If
        On Error GoTo 0

        Dim modeVal As String: modeVal = ""
        Dim planParts() As String
        planParts = Split(plan, "_")
        If UBound(planParts) >= 0 Then
            modeVal = planParts(UBound(planParts))
        End If

        If Len(sessionVal) > 0 And Len(modeVal) > 0 Then
            Dim key As String
            key = sessionVal & "|" & modeVal
            If Not dict.Exists(key) Then
                dict.Add key, bufferBp
            End If
        End If
ContinueLoop:
    Loop
    stream.Close

    Set LoadSlippageOverrides = dict
End Function

Private Function BuildAdjustedPriceFormula(ByVal baseRef As String, ByVal bufferFrac As Double, ByVal isBuy As Boolean) As String
    If Len(baseRef) = 0 Then
        BuildAdjustedPriceFormula = "=" & DQ & DQ
        Exit Function
    End If

    If bufferFrac <= 0 Then
        BuildAdjustedPriceFormula = "=" & baseRef
    Else
        Dim formatted As String
        formatted = Replace(Format$(bufferFrac, "0.################"), ",", ".")
        Dim factorExpr As String
        If isBuy Then
            factorExpr = "(1-" & formatted & ")"
        Else
            factorExpr = "(1+" & formatted & ")"
        End If
        BuildAdjustedPriceFormula = "=IF(" & baseRef & "=" & DQ & DQ & "," & DQ & DQ & "," & baseRef & "*" & factorExpr & ")"
    End If
End Function

Private Function BuildNoteFormula(ByVal sessionRef As String, ByVal modeRef As String, ByVal noteAppend As String) As String
    Dim baseNote As String
    If sessionRef <> "" Or modeRef <> "" Then
        baseNote = "=" & DQ & "session=" & DQ
        If sessionRef <> "" Then
            baseNote = baseNote & " & IF(ISBLANK(" & sessionRef & ")," & DQ & DQ & "," & sessionRef & ")"
        Else
            baseNote = baseNote & " & " & DQ & DQ
        End If
        baseNote = baseNote & " & " & DQ & ";mode=" & DQ
        If modeRef <> "" Then
            baseNote = baseNote & " & IF(ISBLANK(" & modeRef & ")," & DQ & DQ & "," & modeRef & ")"
        Else
            baseNote = baseNote & " & " & DQ & DQ
        End If
    Else
        baseNote = ""
    End If

    If Len(noteAppend) > 0 Then
        If baseNote = "" Then
            BuildNoteFormula = "=" & DQ & noteAppend & DQ
        Else
            BuildNoteFormula = baseNote & " & " & DQ & ";" & DQ & " & " & DQ & noteAppend & DQ
        End If
    Else
        BuildNoteFormula = baseNote
    End If
End Function

Private Function GetSlippageKey(ByVal sessionVal As String, ByVal modeVal As String) As String
    If Len(sessionVal) = 0 Or Len(modeVal) = 0 Then
        GetSlippageKey = ""
    Else
        GetSlippageKey = sessionVal & "|" & modeVal
    End If
End Function

Public Sub ApplyDynamicSignalsV2()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < DASH2_DATA_START Then Exit Sub



    Dim jCol As Long: jCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJValueJP())

    Dim jthCol As Long: jthCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderJThJP()) ' treated as adjusted

    Dim jthBaseCol As Long: jthBaseCol = FindColumn(ws, DASH2_HEADER_ROW, "J_th_base")

    Dim vwapCol As Long: vwapCol = FindColumn(ws, DASH2_HEADER_ROW, "VWAP")

    Dim prevCol As Long: prevCol = FindColumn(ws, DASH2_HEADER_ROW, "PrevClose")

    Dim gapCol As Long: gapCol = FindColumn(ws, DASH2_HEADER_ROW, "Gap_bp")

    Dim corrCol As Long: corrCol = FindColumn(ws, DASH2_HEADER_ROW, "CorrNKY")

    Dim eBuyCol As Long: eBuyCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryBuyPx")

    Dim eSellCol As Long: eSellCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySellPx")

    Dim sideCol As Long: sideCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySide")



    ' Parameter cells on row2 (row1 is labels)

    Dim biasBpRef As String: biasBpRef = ws.Cells(2, 4).Address(False, False, xlR1C1) ' D2

    Dim biasSlopeRef As String: biasSlopeRef = ws.Cells(2, 5).Address(False, False, xlR1C1) ' E2

    Dim gapSlopeRef As String: gapSlopeRef = ws.Cells(2, 6).Address(False, False, xlR1C1)  ' F2

    Dim gapBanRef As String: gapBanRef = ws.Cells(2, 7).Address(False, False, xlR1C1)      ' G2

    Dim corrSlopeRef As String: corrSlopeRef = ws.Cells(2, 12).Address(False, False, xlR1C1) ' L2



    ' J_th adjusted = base + BiasSlope*(Bias_bp/100) + GapSlope*abs(Gap%)+ CorrSlope*CorrNKY*(Bias_bp/100)

    If jthCol > 0 And jthBaseCol > 0 Then

        Dim gapExpr As String: gapExpr = BuildR1C1Ref(gapCol, jthCol)

        Dim corrExpr As String: corrExpr = BuildR1C1Ref(corrCol, jthCol)

        Dim baseRef As String: baseRef = BuildR1C1Ref(jthBaseCol, jthCol)

        Dim jthF As String

        jthF = "=IF(ABS(" & gapExpr & ")/100>" & gapBanRef & "," & DQ & "BAN" & DQ & "," & baseRef & "+(" & biasSlopeRef & ")*" & biasBpRef & "/100+(" & gapSlopeRef & ")*ABS(" & gapExpr & ")/100+(" & corrSlopeRef & ")*" & corrExpr & "*" & biasBpRef & "/100)"

        SetColumnFormula ws, jthCol, lastRow, jthF

    End If



    If eBuyCol > 0 And eSellCol > 0 And vwapCol > 0 And jthCol > 0 Then

        Dim vwapEB As String: vwapEB = BuildR1C1Ref(vwapCol, eBuyCol)

        Dim vwapES As String: vwapES = BuildR1C1Ref(vwapCol, eSellCol)

        Dim jthEB As String: jthEB = BuildR1C1Ref(jthCol, eBuyCol)

        Dim baseB As String: baseB = "IF(" & vwapEB & "=" & DQ & DQ & ",RC[" & (prevCol - eBuyCol) & "]," & vwapEB & ")"

        Dim baseS As String: baseS = "IF(" & vwapES & "=" & DQ & DQ & ",RC[" & (prevCol - eSellCol) & "]," & vwapES & ")"

        Dim k As String: k = "0.001"

        Dim eBuyF As String

        Dim eSellF As String

        eBuyF = "=IF(OR(" & baseB & "=" & DQ & DQ & "," & jthEB & "=" & DQ & DQ & "," & jthEB & "=" & DQ & "BAN" & DQ & ")," & DQ & DQ & ",(" & baseB & ")-" & k & "*ABS(" & jthEB & ")*" & baseB & ")"

        eSellF = "=IF(OR(" & baseS & "=" & DQ & DQ & "," & jthEB & "=" & DQ & DQ & "," & jthEB & "=" & DQ & "BAN" & DQ & ")," & DQ & DQ & ",(" & baseS & ")+" & k & "*ABS(" & jthEB & ")*" & baseS & ")"

        SetColumnFormula ws, eBuyCol, lastRow, eBuyF

        SetColumnFormula ws, eSellCol, lastRow, eSellF

    End If



    If sideCol > 0 And jCol > 0 Then

        Dim jSide As String: jSide = BuildR1C1Ref(jCol, sideCol)

        SetColumnFormula ws, sideCol, lastRow, "=IF(" & jSide & "<0," & DQ & "BUY" & DQ & ",IF(" & jSide & ">0," & DQ & "SELL" & DQ & "," & DQ & DQ & "))"

    End If

End Sub



Public Sub PreplaceOrdersV2()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim selCol As Long: selCol = FindColumn(ws, DASH2_HEADER_ROW, "Selected")
    Dim eBuyCol As Long: eBuyCol = FindColumn(ws, DASH2_HEADER_ROW, "EntryBuyPx")
    Dim eSellCol As Long: eSellCol = FindColumn(ws, DASH2_HEADER_ROW, "EntrySellPx")
    Dim qtyCol As Long: qtyCol = FindColumn(ws, DASH2_HEADER_ROW, "OrderQtyPlan")
    Dim tpCol As Long: tpCol = FindColumn(ws, DASH2_HEADER_ROW, "TP_price")
    Dim slCol As Long: slCol = FindColumn(ws, DASH2_HEADER_ROW, "SL_price")
    Dim modeCol As Long: modeCol = FindColumn(ws, DASH2_HEADER_ROW, "SignalMode")
    Dim sessionCol As Long: sessionCol = FindColumn(ws, DASH2_HEADER_ROW, "session")
    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH2_HEADER_ROW, HeaderTickerJP())

    If selCol = 0 Or tickerCol = 0 Then Exit Sub

    Dim sh As Worksheet
    Set sh = EnsureOrdersSheet(ws)

    Dim slipDict As Object
    Set slipDict = LoadSlippageOverrides()

    Dim lastOrder As Long: lastOrder = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim idx As Long
    For idx = lastOrder To 2 Step -1
        If LCase$(CStr(sh.Cells(idx, 6).Value)) = "preplace" Then
            sh.Rows(idx).Delete
        End If
    Next idx

    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = DASH2_DATA_START To lastR
        If ws.Cells(r, selCol).Value = 1 Then
            Dim hasBuy As Boolean: hasBuy = False
            Dim hasSell As Boolean: hasSell = False
            Dim tmpVal As Variant
            If eBuyCol > 0 Then
                hasBuy = ws.Cells(r, eBuyCol).HasFormula
                If Not hasBuy Then
                    tmpVal = ws.Cells(r, eBuyCol).Value
                    hasBuy = IsNumeric(tmpVal) And tmpVal > 0
                End If
            End If
            If eSellCol > 0 Then
                hasSell = ws.Cells(r, eSellCol).HasFormula
                If Not hasSell Then
                    tmpVal = ws.Cells(r, eSellCol).Value
                    hasSell = IsNumeric(tmpVal) And tmpVal > 0
                End If
            End If
            If hasBuy Or hasSell Then
                Dim modeVal As String: modeVal = ""
                Dim sessionVal As String: sessionVal = ""
                If modeCol > 0 Then modeVal = CStr(ws.Cells(r, modeCol).Value)
                If sessionCol > 0 Then sessionVal = CStr(ws.Cells(r, sessionCol).Value)
                Dim noteExtra As String: noteExtra = ""
                Dim bufferFrac As Double: bufferFrac = 0
                Dim key As String: key = GetSlippageKey(sessionVal, modeVal)
                If Len(key) > 0 Then
                    If Not slipDict Is Nothing Then
                        If slipDict.Exists(key) Then
                            bufferFrac = CDbl(slipDict(key)) / 10000#
                            If bufferFrac < 0 Then bufferFrac = 0
                            If bufferFrac > 0.003 Then bufferFrac = 0.003
                            noteExtra = "buffer_bp=" & Format$(slipDict(key), "0.0")
                        End If
                    End If
                End If
                LogPreOrder ws, sh, r, hasBuy, hasSell, eBuyCol, eSellCol, qtyCol, tpCol, slCol, modeCol, sessionCol, tickerCol, bufferFrac, noteExtra
            End If
        End If
    Next r

End Sub



Private Function EnsureOrdersSheet(ByVal host As Worksheet) As Worksheet

    Dim sh As Worksheet

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("Orders")
    On Error GoTo 0

    If sh Is Nothing Then
        Dim anchor As Worksheet
        If host Is Nothing Then
            Set anchor = ThisWorkbook.Worksheets(1)
        Else
            Set anchor = host
        End If
        Set sh = ThisWorkbook.Worksheets.Add(After:=anchor)
        sh.Name = "Orders"
        sh.Range("A1:K1").Value = Array("ts", "ticker", "side", "price", "qty", "mode", "status", "note", "tp", "sl", "trail")
    End If

    Set EnsureOrdersSheet = sh

End Function


Private Sub LogPreOrder(ByVal ws As Worksheet, ByVal orderSheet As Worksheet, ByVal rowIndex As Long, _
    ByVal includeBuy As Boolean, ByVal includeSell As Boolean, _
    ByVal eBuyCol As Long, ByVal eSellCol As Long, ByVal qtyCol As Long, _
    ByVal tpCol As Long, ByVal slCol As Long, ByVal modeCol As Long, _
    ByVal sessionCol As Long, ByVal tickerCol As Long, ByVal bufferFrac As Double, ByVal noteExtra As String)

    Dim sh As Worksheet
    Set sh = orderSheet
    If sh Is Nothing Then
        Set sh = EnsureOrdersSheet(ws)
    End If

    Dim sheetToken As String
    sheetToken = "'" & Replace(ws.Name, "'", "''") & "'!"

    Dim tickerRef As String
    If tickerCol > 0 Then
        tickerRef = sheetToken & ws.Cells(rowIndex, tickerCol).Address(True, True, xlA1)
    Else
        tickerRef = ""
    End If

    Dim qtyRef As String
    If qtyCol > 0 Then qtyRef = sheetToken & ws.Cells(rowIndex, qtyCol).Address(True, True, xlA1) Else qtyRef = ""

    Dim tpRef As String
    If tpCol > 0 Then tpRef = sheetToken & ws.Cells(rowIndex, tpCol).Address(True, True, xlA1) Else tpRef = ""

    Dim slRef As String
    If slCol > 0 Then slRef = sheetToken & ws.Cells(rowIndex, slCol).Address(True, True, xlA1) Else slRef = ""

    Dim sessionRef As String
    If sessionCol > 0 Then sessionRef = sheetToken & ws.Cells(rowIndex, sessionCol).Address(True, True, xlA1) Else sessionRef = ""

    Dim modeRef As String
    If modeCol > 0 Then modeRef = sheetToken & ws.Cells(rowIndex, modeCol).Address(True, True, xlA1) Else modeRef = ""

    Dim noteFormula As String
    noteFormula = BuildNoteFormula(sessionRef, modeRef, noteExtra)

    Dim nowStr As String: nowStr = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim nextRow As Long: nextRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1

    If includeBuy Then
        sh.Cells(nextRow, 1).Value = nowStr
        If tickerRef <> "" Then
            sh.Cells(nextRow, 2).Formula = "=" & tickerRef
        Else
            sh.Cells(nextRow, 2).Value = ""
        End If
        sh.Cells(nextRow, 3).Value = "BUY"
        If eBuyCol > 0 Then
            Dim buyRef As String
            buyRef = sheetToken & ws.Cells(rowIndex, eBuyCol).Address(True, True, xlA1)
            sh.Cells(nextRow, 4).Formula = BuildAdjustedPriceFormula(buyRef, bufferFrac, True)
        Else
            sh.Cells(nextRow, 4).Value = ""
        End If
        If qtyRef <> "" Then
            sh.Cells(nextRow, 5).Formula = "=" & qtyRef
        Else
            sh.Cells(nextRow, 5).Value = 0
        End If
        sh.Cells(nextRow, 6).Value = "preplace"
        sh.Cells(nextRow, 7).Value = "PENDING"
        If noteFormula <> "" Then
            sh.Cells(nextRow, 8).Formula = noteFormula
        Else
            sh.Cells(nextRow, 8).Value = ""
        End If
        If tpRef <> "" Then
            sh.Cells(nextRow, 9).Formula = "=" & tpRef
        Else
            sh.Cells(nextRow, 9).Value = ""
        End If
        If slRef <> "" Then
            sh.Cells(nextRow, 10).Formula = "=" & slRef
        Else
            sh.Cells(nextRow, 10).Value = ""
        End If
        sh.Cells(nextRow, 11).Value = ""
        nextRow = nextRow + 1
    End If

    If includeSell Then
        sh.Cells(nextRow, 1).Value = nowStr
        If tickerRef <> "" Then
            sh.Cells(nextRow, 2).Formula = "=" & tickerRef
        Else
            sh.Cells(nextRow, 2).Value = ""
        End If
        sh.Cells(nextRow, 3).Value = "SELL"
        If eSellCol > 0 Then
            Dim sellRef As String
            sellRef = sheetToken & ws.Cells(rowIndex, eSellCol).Address(True, True, xlA1)
            sh.Cells(nextRow, 4).Formula = BuildAdjustedPriceFormula(sellRef, bufferFrac, False)
        Else
            sh.Cells(nextRow, 4).Value = ""
        End If
        If qtyRef <> "" Then
            sh.Cells(nextRow, 5).Formula = "=" & qtyRef
        Else
            sh.Cells(nextRow, 5).Value = 0
        End If
        sh.Cells(nextRow, 6).Value = "preplace"
        sh.Cells(nextRow, 7).Value = "PENDING"
        If noteFormula <> "" Then
            sh.Cells(nextRow, 8).Formula = noteFormula
        Else
            sh.Cells(nextRow, 8).Value = ""
        End If
        If tpRef <> "" Then
            sh.Cells(nextRow, 9).Formula = "=" & tpRef
        Else
            sh.Cells(nextRow, 9).Value = ""
        End If
        If slRef <> "" Then
            sh.Cells(nextRow, 10).Formula = "=" & slRef
        Else
            sh.Cells(nextRow, 10).Value = ""
        End If
        sh.Cells(nextRow, 11).Value = ""
    End If

End Sub



Public Function PlaceOrderRSS(ByVal ticker As String, ByVal side As String, ByVal price As Double, ByVal qty As Long) As Boolean

    On Error GoTo Fail

    Dim evalStr As String

    evalStr = "RssOrder(\"" & ticker & " \ ",\"" & side & " \ "," & CStr(price) & "," & CStr(qty) & ")"

    Dim result As Variant

    result = Application.Evaluate("=" & evalStr)

    If IsError(result) Then GoTo Fail

    PlaceOrderRSS = True

    Exit Function

Fail:

    PlaceOrderRSS = False

End Function



' ----------------------------------------------------------------------------

' UI: Buttons/Status/Labels/Order

' ----------------------------------------------------------------------------

Public Sub SetupDashboardUIV2()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    ' Move status banner to the right to avoid overlapping params

    With ws.Range("N1:U2")

        .Merge

        .HorizontalAlignment = xlCenter

        .VerticalAlignment = xlCenter

        .Font.Bold = True

        .Font.Size = 18

        .Interior.Color = RGB(230, 230, 230)

        .Value = "IDLE"

        .name = "RunStatusV2"

    End With



    Dim shp As Shape

    For Each shp In ws.Shapes

        If shp.name Like "btn_*" Then shp.Delete

    Next shp



    ' Buttons (swap: Live on left, Demo on right)

    CreateButton ws, "btn_live_start", "Live Start", 3, 14, "AutoTraderAdvanced.StartLiveV2"    ' N3

    CreateButton ws, "btn_live_stop", "Live Stop", 3, 18, "AutoTraderAdvanced.StopLiveV2"      ' R3

    CreateButton ws, "btn_demo_start", "Demo Start", 3, 22, "AutoTraderAdvanced.StartDemoV2"    ' V3

    CreateButton ws, "btn_demo_stop", "Demo Stop", 3, 26, "AutoTraderAdvanced.StopDemoV2"      ' Z3

    CreateButton ws, "btn_import", "Import Candidates", 3, 30, "AutoTraderAdvanced.ImportCandidatesV2"      ' AD3



    ApplyJapaneseLabelsV2 ws

    ReorderHeadersV2 ws

End Sub



Private Sub CreateButton(ByVal ws As Worksheet, ByVal name As String, ByVal caption As String, ByVal rowTop As Long, ByVal colLeft As Long, ByVal onAction As String)

    Dim leftPos As Double: leftPos = ws.Cells(rowTop, colLeft).Left

    Dim topPos As Double: topPos = ws.Cells(rowTop, colLeft).Top

    Dim width As Double: width = ws.Range(ws.Cells(rowTop, colLeft), ws.Cells(rowTop, colLeft + 3)).width - 6

    Dim height As Double: height = ws.Rows(rowTop).height - 2

    Dim btn As Shape

    Set btn = ws.Shapes.AddShape(MsoShapeRectangle, leftPos, topPos, width, height)

    On Error Resume Next

    btn.name = name

    On Error GoTo 0

    btn.TextFrame.Characters.Text = caption

    btn.TextFrame.HorizontalAlignment = MsoAlignCenter

    btn.TextFrame.VerticalAlignment = MsoAlignCenter

    btn.onAction = onAction

End Sub



Private Sub ApplyJapaneseLabelsV2(ByVal ws As Worksheet)

    Dim labels As Variant

    labels = Array("Ticker", "Name", "J_th_base", "J_th", "J", "PrevClose", "VWAP", "OrderQtyPlan", "Selected", _

                   "EntryBuyPx", "EntrySellPx", "EntrySide", "EntryStatus", "TP_price", "SL_price", "StopTrail", "SettleStatus", "BestBid", "BestAsk", "Gap_bp", "CorrNKY")

    Dim i As Long

    For i = 1 To 20

        ws.Cells(4, i).Value = labels(i - 1)

        On Error Resume Next

        ws.Cells(4, i).AddComment "Header description"

        On Error GoTo 0

    Next i

End Sub



Private Sub ReorderHeadersV2(ByVal ws As Worksheet)

    Dim order As Variant

    order = Array("Ticker", "Name", "J_th_base", "J_th", "J", "PrevClose", "VWAP", "OrderQtyPlan", "Selected", "EntryBuyPx", "EntrySellPx", "EntrySide", "EntryStatus", "TP_price", "SL_price", "StopTrail", "SettleStatus", "BestBid", "BestAsk", "Gap_bp", "CorrNKY")

    Dim c As Long

    For c = 1 To UBound(order) + 1

        ws.Cells(DASH2_HEADER_ROW, c).Value = CStr(order(c - 1))

    Next c

End Sub



Public Sub StartDemoV2(): UpdateStatusV2 "DEMO_RUNNING": End Sub

Public Sub StopDemoV2():  UpdateStatusV2 "IDLE": End Sub

Public Sub StartLiveV2(): UpdateStatusV2 "LIVE_RUNNING": End Sub

Public Sub StopLiveV2():  UpdateStatusV2 "IDLE": End Sub



Private Sub UpdateStatusV2(ByVal mode As String)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    Dim statusArea As Range
    Set statusArea = ws.Range("A3:B3")

    On Error Resume Next
    statusArea.UnMerge
    On Error GoTo 0

    On Error Resume Next
    ws.Parent.Names.Add Name:="RunStatusV2", RefersTo:=statusArea
    On Error GoTo 0

    statusArea.ClearContents

    With statusArea
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 16

        Select Case mode
            Case "DEMO_RUNNING": .Interior.Color = RGB(220, 240, 255)
            Case "LIVE_RUNNING": .Interior.Color = RGB(255, 230, 230)
            Case Else: .Interior.Color = RGB(230, 230, 230)
        End Select
    End With

    statusArea.Cells(1, 1).Value = mode

End Sub



Public Sub ImportCandidatesV2()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)

    Dim path As String

    path = ThisWorkbook.path & "\output\excel\candidates_nextday.csv"

    On Error Resume Next

    If Len(Dir$(path)) = 0 Then

        Dim dt As String: dt = Format$(Date, "yyyymmdd")

        path = ThisWorkbook.path & "\output\excel\weekly_candidates_" & dt & ".csv"

    End If

    On Error GoTo 0

    If Len(Dir$(path)) = 0 Then Exit Sub

    Dim f As Integer: f = FreeFile

    Open path For Input As #f

    Dim line As String

    Dim r As Long: r = DASH2_DATA_START

    Dim selCol As Long: selCol = FindColumn(ws, DASH2_HEADER_ROW, "Selected")

    Dim jtbCol As Long: jtbCol = FindColumn(ws, DASH2_HEADER_ROW, "J_th_base")

    Dim colPf As Long: colPf = FindColumn(ws, DASH2_HEADER_ROW, "ForwardPfEff")

    Dim colCi As Long: colCi = FindColumn(ws, DASH2_HEADER_ROW, "WinCiLow")

    Dim colTr As Long: colTr = FindColumn(ws, DASH2_HEADER_ROW, "ForwardTrades")

    Dim colExp As Long: colExp = FindColumn(ws, DASH2_HEADER_ROW, "ExpBp")

    Dim colAtr As Long: colAtr = FindColumn(ws, DASH2_HEADER_ROW, "ATR_n")

    Dim colTpk As Long: colTpk = FindColumn(ws, DASH2_HEADER_ROW, "TPk")

    Dim colSlk As Long: colSlk = FindColumn(ws, DASH2_HEADER_ROW, "SLk")

    Dim colMode As Long: colMode = FindColumn(ws, DASH2_HEADER_ROW, "SignalMode")

    Dim colSession As Long: colSession = FindColumn(ws, DASH2_HEADER_ROW, "session")

    Dim colPlan As Long: colPlan = FindColumn(ws, DASH2_HEADER_ROW, "plan_tag")

    Dim maxExisting As Long: maxExisting = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim first As Boolean: first = True

    Dim idxTicker As Long, idxJtb As Long, idxPf As Long, idxCi As Long, idxTrades As Long, idxExpBp As Long

    Dim idxAtr As Long, idxTpk As Long, idxSlk As Long, idxMode As Long, idxSession As Long, idxPlan As Long

    idxTicker = -1: idxJtb = -1: idxPf = -1: idxCi = -1: idxTrades = -1: idxExpBp = -1

    idxAtr = -1: idxTpk = -1: idxSlk = -1: idxMode = -1: idxSession = -1: idxPlan = -1

    Dim hdr As Variant
    Do While Not EOF(f)

        Line Input #f, line

        If first Then

            hdr = ParseCsvLine(line)
            Dim i As Long

            For i = LBound(hdr) To UBound(hdr)

                Dim h As String: h = LCase$(Trim$(hdr(i)))
                h = Application.WorksheetFunction.Clean(h)
                Do While Len(h) > 0
                    Dim chCode As Long: chCode = AscW(Left$(h, 1))
                    If chCode >= 48 And chCode <= 122 Then Exit Do
                    h = Mid$(h, 2)
                Loop
                If Len(h) > 0 Then
                    If AscW(Left$(h, 1)) = &HFEFF Then
                        h = Mid$(h, 2)
                    End If
                End If
                If h = "ticker" Then idxTicker = i
                If h = "code" And idxTicker = -1 Then idxTicker = i

                If h = "j_th_base" Or h = "j_th" Then idxJtb = i

                If h = "forward_pf_eff" Or h = "pf" Then idxPf = i

                If h = "forward_win_ci_low" Or h = "win_ci_low" Then idxCi = i

                If h = "forward_trades" Or h = "trades" Then idxTrades = i

                If h = "exp_bp" Or h = "expected_bp" Or h = "forward_exp_bp" Or h = "forward_exp_boot_mean" Then idxExpBp = i
                If h = "atr_n" Then idxAtr = i

                If h = "tpk" Then idxTpk = i

                If h = "slk" Then idxSlk = i

                If h = "signalmode" Or h = "signal_mode" Then idxMode = i

                If h = "session" Then idxSession = i

                If h = "plan_tag" Then idxPlan = i

            Next i

            first = False

        Else

            Dim parts As Variant: parts = ParseCsvLine(line)
            If idxTicker >= 0 And idxTicker <= UBound(parts) Then

                Dim tkr As String: tkr = Trim$(parts(idxTicker))
                If Len(tkr) > 1 And Left$(tkr, 1) = """" And Right$(tkr, 1) = """" Then
                    tkr = Mid$(tkr, 2, Len(tkr) - 2)
                End If
                tkr = Replace$(tkr, """", "")
                If Len(tkr) > 0 Then

                    ws.Cells(r, 1).Value = tkr

                    If selCol > 0 Then ws.Cells(r, selCol).Value = 1

                    If idxJtb >= 0 And idxJtb <= UBound(parts) And jtbCol > 0 Then

                        If Len(Trim$(parts(idxJtb))) > 0 Then ws.Cells(r, jtbCol).Value = Trim$(parts(idxJtb))

                    End If

                    If idxPf >= 0 And idxPf <= UBound(parts) And colPf > 0 Then ws.Cells(r, colPf).Value = Trim$(parts(idxPf))

                    If idxCi >= 0 And idxCi <= UBound(parts) And colCi > 0 Then ws.Cells(r, colCi).Value = Trim$(parts(idxCi))

                    If idxTrades >= 0 And idxTrades <= UBound(parts) And colTr > 0 Then ws.Cells(r, colTr).Value = Trim$(parts(idxTrades))

                    If idxExpBp >= 0 And idxExpBp <= UBound(parts) And colExp > 0 Then ws.Cells(r, colExp).Value = Trim$(parts(idxExpBp))

                    If idxAtr >= 0 And idxAtr <= UBound(parts) And colAtr > 0 Then

                        If Len(Trim$(parts(idxAtr))) > 0 Then

                            ws.Cells(r, colAtr).Value = Trim$(parts(idxAtr))

                        ElseIf ws.Cells(r, colAtr).Value = "" Then

                            ws.Cells(r, colAtr).Value = 2

                        End If

                    ElseIf colAtr > 0 And ws.Cells(r, colAtr).Value = "" Then

                        ws.Cells(r, colAtr).Value = 2

                    End If

                    If idxTpk >= 0 And idxTpk <= UBound(parts) And colTpk > 0 Then ws.Cells(r, colTpk).Value = Trim$(parts(idxTpk))

                    If idxSlk >= 0 And idxSlk <= UBound(parts) And colSlk > 0 Then ws.Cells(r, colSlk).Value = Trim$(parts(idxSlk))

                    If idxMode >= 0 And idxMode <= UBound(parts) And colMode > 0 Then ws.Cells(r, colMode).Value = Trim$(parts(idxMode))

                    If idxSession >= 0 And idxSession <= UBound(parts) And colSession > 0 Then ws.Cells(r, colSession).Value = Trim$(parts(idxSession))

                    If idxPlan >= 0 And idxPlan <= UBound(parts) And colPlan > 0 Then ws.Cells(r, colPlan).Value = Trim$(parts(idxPlan))

                    r = r + 1

                End If

            End If

        End If

    Loop

    Close #f



    Dim clearRow As Long

    If maxExisting >= DASH2_DATA_START And r <= maxExisting Then

        For clearRow = r To maxExisting

            ws.Cells(clearRow, 1).ClearContents

            If selCol > 0 Then ws.Cells(clearRow, selCol).ClearContents

            If jtbCol > 0 Then ws.Cells(clearRow, jtbCol).ClearContents

            If colPf > 0 Then ws.Cells(clearRow, colPf).ClearContents

            If colCi > 0 Then ws.Cells(clearRow, colCi).ClearContents

            If colTr > 0 Then ws.Cells(clearRow, colTr).ClearContents

            If colExp > 0 Then ws.Cells(clearRow, colExp).ClearContents

            If colAtr > 0 Then ws.Cells(clearRow, colAtr).ClearContents

            If colTpk > 0 Then ws.Cells(clearRow, colTpk).ClearContents

            If colSlk > 0 Then ws.Cells(clearRow, colSlk).ClearContents

            If colMode > 0 Then ws.Cells(clearRow, colMode).ClearContents

            If colSession > 0 Then ws.Cells(clearRow, colSession).ClearContents

            If colPlan > 0 Then ws.Cells(clearRow, colPlan).ClearContents

        Next clearRow

    End If

    ws.Calculate
    ApplyDynamicSignalsV2
    PreplaceOrdersV2

End Sub







