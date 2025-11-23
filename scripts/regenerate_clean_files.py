import os

def write_file(path, content):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Wrote {path}")

clean_ensure_params = '''Private Sub EnsureParamFormulas(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells(2, 1).value = "N225"
    ws.Cells(2, 4).value = "TOPX"
    ws.Cells(2, 2).Formula = "=IF(A2="""","""",IFERROR(RssIndexMarket(A2,""現在値""),""""))"
    ws.Cells(2, 3).Formula = "=IF(A2="""","""",IFERROR(RssIndexMarket(A2,""前日比率""),""""))"
    ws.Cells(2, 5).Formula = "=IF(D2="""","""",IFERROR(RssIndexMarket(D2,""現在値""),""""))"
    ws.Cells(2, 6).Formula = "=IF(D2="""","""",IFERROR(RssIndexMarket(D2,""前日比率""),""""))"
    
    ' Split params into two arrays to avoid line continuation limit
    Dim p1 As Variant, p2 As Variant
    p1 = Array( _
        Array(1, "指標コード(日経平均)", "楽天RSSに渡す指標コード。通常はN225固定です"), _
        Array(2, "日経平均 現在値", ""), _
        Array(3, "日経平均 前日比率", ""), _
        Array(4, "指標コード(TOPIX)", "楽天RSSに渡すTOPIXコード。通常はTOPXです"), _
        Array(5, "TOPIX 現在値", ""), _
        Array(6, "TOPIX 前日比率", ""), _
        Array(7, "バイアス閾値(bp)", "J補正をBAN扱いにする絶対値閾値"), _
        Array(8, "Bias補正係数", "日経平均との方向差で加算する係数"), _
        Array(9, "Gap補正係数", "ギャップ量に応じた補正係数"), _
        Array(10, "Gap BAN 閾値(%)", "この割合を超えるギャップは自動BAN"), _
        Array(11, "取引停止分数", "AutoTrader再開までのクールダウン（分）"), _
        Array(12, "TP/J (全体)", "") _
    )
    p2 = Array( _
        Array(13, "SL/J (全体)", ""), _
        Array(14, "Trail/J (全体)", ""), _
        Array(15, "相関補正係数", "NKY/TOPIX相関でJ_thを補正する係数"), _
        Array(16, "銘柄別予算(円)", ""), _
        Array(17, "ロットサイズ", ""), _
        Array(18, "NKY日足トレンド", ""), _
        Array(19, "NKY窓トレンド", ""), _
        Array(20, "NKY許容サイド", ""), _
        Array(21, "TOPIX日足トレンド", ""), _
        Array(22, "TOPIX窓トレンド", ""), _
        Array(23, "TOPIX許容サイド", ""), _
        Array(24, "注文訂正閾値(tick)", "Preplace/決済注文を訂正する最小乖離幅") _
    )
    
    Dim headerInfo As Variant
    Dim i As Long
    For i = LBound(p1) To UBound(p1)
        headerInfo = p1(i)
        ws.Cells(1, headerInfo(0)).value = headerInfo(1)
        If UBound(headerInfo) >= 2 Then
            If Len(headerInfo(2)) > 0 Then SetHeaderComment ws, CLng(headerInfo(0)), CStr(headerInfo(2))
        End If
    Next i
    For i = LBound(p2) To UBound(p2)
        headerInfo = p2(i)
        ws.Cells(1, headerInfo(0)).value = headerInfo(1)
        If UBound(headerInfo) >= 2 Then
            If Len(headerInfo(2)) > 0 Then SetHeaderComment ws, CLng(headerInfo(0)), CStr(headerInfo(2))
        End If
    Next i
    On Error GoTo 0
End Sub'''

clean_setup_ui = '''Sub SetupDashboardUIV2()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH2_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Recreate Buttons (Standard V2 Layout)
    CreateButton ws, "btn_live_stop", "Live Stop", 3, 8, "AutoTraderAdvanced.StopLive"
    CreateButton ws, "btn_demo_start", "Demo Start", 3, 10, "AutoTraderAdvanced.StartDemo"
    CreateButton ws, "btn_demo_stop", "Demo Stop", 3, 12, "AutoTraderAdvanced.StopDemo"
    CreateButton ws, "btn_import", "Import Candidates", 3, 14, "AutoTraderAdvanced.ImportCandidates"
    CreateButton ws, "btn_recalc", "方向再計算", 3, 16, "AutoTraderAdvanced.RecalcDirection"
    CreateButton ws, "btn_clear_bb", "BBブロック解除", 3, 18, "AutoTraderAdvanced.ClearBBBlocks"
    
    ApplyJapaneseLabelsV2 ws
    ReorderHeadersV2 ws
    EnsureParamFormulas ws
    SetupHeartbeatCell ws
    UpdateTrendIndicators ws
    
    With ws.Range("A3")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
    End With
End Sub

Private Sub CreateButton(ByVal ws As Worksheet, ByVal name As String, ByVal caption As String, ByVal rowTop As Long, ByVal colLeft As Long, ByVal macroName As String)
    Dim btn As Button
    Dim rng As Range
    Dim leftPos As Double, topPos As Double, width As Double, height As Double
    
    On Error Resume Next
    ws.Buttons(name).Delete
    On Error GoTo 0
    
    Set rng = ws.Cells(rowTop, colLeft)
    leftPos = rng.Left
    topPos = rng.Top
    width = rng.Width * 2
    height = rng.Height * 1.5
    
    Set btn = ws.Buttons.Add(leftPos, topPos, width, height)
    With btn
        .name = name
        .caption = caption
        .OnAction = macroName
        .Font.Size = 10
        .Font.Bold = True
    End With
End Sub'''

clean_log_preorder = '''Private Sub LogPreOrder(ByVal ws As Worksheet, ByVal Sh As Worksheet, ByVal rowIndex As Long, _
    ByVal includeBuy As Boolean, ByVal includeSell As Boolean, _
    ByVal eBuyCol As Long, ByVal eSellCol As Long, ByVal qtyCol As Long, _
    ByVal tpCol As Long, ByVal slCol As Long, ByVal modeCol As Long, _
    ByVal sessionCol As Long, ByVal tickerCol As Long, _
    ByVal bufferFrac As Double, ByVal noteExtra As String)

    Dim nextRow As Long
    nextRow = Sh.Cells(Sh.Rows.Count, 1).End(xlUp).Row + 1
    
    Dim sheetToken As String
    sheetToken = "'" & ws.name & "'!"
    
    Dim tickerRef As String
    tickerRef = sheetToken & ws.Cells(rowIndex, tickerCol).address(True, True)
    
    Dim qtyRef As String
    qtyRef = sheetToken & ws.Cells(rowIndex, qtyCol).address(True, True)
    
    Dim tpRef As String
    If tpCol > 0 Then tpRef = sheetToken & ws.Cells(rowIndex, tpCol).address(True, True)
    Dim slRef As String
    If slCol > 0 Then slRef = sheetToken & ws.Cells(rowIndex, slCol).address(True, True)
    
    Dim noteFormula As String
    noteFormula = ""
    If Len(noteExtra) > 0 Then
        noteFormula = "=""" & noteExtra & """"
    End If

    If includeBuy Then
        Dim buyRef As String
        buyRef = sheetToken & ws.Cells(rowIndex, eBuyCol).address(True, True)
        
        Sh.Cells(nextRow, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        Sh.Cells(nextRow, 2).Formula = "=" & tickerRef
        Sh.Cells(nextRow, 3).value = "BUY"
        Sh.Cells(nextRow, 4).Formula = "=" & buyRef
        Sh.Cells(nextRow, 5).Formula = "=" & qtyRef
        Sh.Cells(nextRow, 6).value = "PREPLACE"
        Sh.Cells(nextRow, 7).value = "PENDING"
        If noteFormula <> "" Then
            Sh.Cells(nextRow, 8).Formula = noteFormula
        Else
            Sh.Cells(nextRow, 8).value = ""
        End If
        
        If tpRef <> "" Then Sh.Cells(nextRow, 9).Formula = "=" & tpRef
        If slRef <> "" Then Sh.Cells(nextRow, 10).Formula = "=" & slRef
        
        nextRow = nextRow + 1
    End If

    If includeSell Then
        Dim sellRef As String
        sellRef = sheetToken & ws.Cells(rowIndex, eSellCol).address(True, True)
        
        Sh.Cells(nextRow, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        Sh.Cells(nextRow, 2).Formula = "=" & tickerRef
        Sh.Cells(nextRow, 3).value = "SELL"
        Sh.Cells(nextRow, 4).Formula = "=" & sellRef
        Sh.Cells(nextRow, 5).Formula = "=" & qtyRef
        Sh.Cells(nextRow, 6).value = "PREPLACE"
        Sh.Cells(nextRow, 7).value = "PENDING"
        If noteFormula <> "" Then
            Sh.Cells(nextRow, 8).Formula = noteFormula
        Else
            Sh.Cells(nextRow, 8).value = ""
        End If
        
        If tpRef <> "" Then Sh.Cells(nextRow, 9).Formula = "=" & tpRef
        If slRef <> "" Then Sh.Cells(nextRow, 10).Formula = "=" & slRef
        
        nextRow = nextRow + 1
    End If
End Sub'''

write_file(r"c:\AI\asagake\scripts\clean_ensure_params.txt", clean_ensure_params)
write_file(r"c:\AI\asagake\scripts\clean_setup.txt", clean_setup_ui)
write_file(r"c:\AI\asagake\scripts\clean_log_preorder.txt", clean_log_preorder)
