Attribute VB_Name = "AutoTrader"
Option Explicit

Private Const SHEET_DASHBOARD As String = "NewDashboard"
Private Const SHEET_CANDIDATES As String = "Candidates"
Private Const SHEET_ORDERS As String = "Orders"
Private Const CANDIDATES_REL_PATH As String = "\output\excel\candidates_nextday.csv"

Private Const DASH_HEADER_ROW As Long = 5
Private Const DASH_DATA_START As Long = 6
Private Const DASH_STATUS_CELL As String = "B2"
Private Const DASH_MAX_ORDERS_CELL As String = "B3"
Private Const DASH_SESSION_START_CELL As String = "B4"
Private Const DASH_SESSION_END_CELL As String = "B5"
Private Const DASH_SELECTED_DEFAULT_CELL As String = "B6"
Private Const DASH_REENTRY_CELL As String = "B7"
Private Const DASH_HARDSTOP_CELL As String = "B8"
Private Const DASH_LIVE_CELL As String = "B9"
Private Const DASH_ORDERMACRO_CELL As String = "B10"
Private Const DASH_CLOSE_TIME_CELL As String = "B11"
Private Const DASH_QTY_CELL As String = "B12"
Private Const DASH_TIF_CELL As String = "B13"
Private Const DASH_BUDGET_CELL As String = "B14"
Private Const DASH_LOT_STEP_CELL As String = "B15"
Private Const DASH_SLIP_BP_CELL As String = "B16"
' 隴・ｽｰ髫募臆・ｨ・ｭ陞ｳ螢ｹ縺晉ｹ晢ｽｫ
Private Const DASH_TRACE_CELL As String = "B17"
Private Const DASH_CANCEL_AT_END_CELL As String = "B18"
Private Const DASH_ENTRY_GRACE_MIN_CELL As String = "B19"
Private Const DASH_OCO_UPDATE_SEC_CELL As String = "B20"
Private Const DEFAULT_MAX_ORDERS As Long = 20
Private Const DEFAULT_SESSION_START As String = "09:00"
Private Const DEFAULT_SESSION_END As String = "09:15"
Private Const DEFAULT_SELECTED_DEFAULT As Long = 1
Private Const DEFAULT_REENTRY As Long = 0
Private Const DEFAULT_HARDSTOP As Long = 0
Private Const DEFAULT_LIVE As Long = 0
Private Const DEFAULT_CLOSE_TIME As String = "14:59:30"
Private Const DEFAULT_ORDER_QTY As Long = 100
Private Const DEFAULT_MAX_BUDGET As Double = 10000000#
Private Const DEFAULT_LOT_STEP As Long = 100
Private Const DEFAULT_SLIP_BP As Double = 30#
Private Const DASH_FORMULA_ROWS As Long = 400
Private Const DASH_MIN_JTH As Double = 0.6
Private Const RISK_PER_TICKER_FRAC As Double = 0.06
Private Const RISK_TOTAL_FRAC As Double = 0.3
Private Const SESSION_MODE_JONLY_RATIO As Double = 0.6
Private Const SESSION_MODE_JCROSS_RATIO As Double = 0.4
Private Const ROW_SCORE_TRADE_TARGET As Double = 40#
Private Const ROW_SCORE_EXP_TARGET As Double = 15#
Private Const ROW_SCORE_DD_K As Double = 2#
Private Const ROW_SCORE_PF_CAP_GROWTH As Double = 2#

Private prevJ As Object
Private AutoTimer As Date
Private isRunning As Boolean
Private tradeDate As Date
Private orderCount As Long
Private setupInitialized As Boolean
Private sessionRiskUsed As Object
Private tickerRiskUsed As Object
Private totalRiskUsed As Double
Private closeTimer As Date
Private closeTimerScheduled As Boolean

Private Function RealTimeHeaderList() As Variant
    RealTimeHeaderList = Array("Ticker", HeaderNameJP(), HeaderJValueJP(), HeaderJGapJP(), HeaderSignalStatusJP(), HeaderSignalKindJP(), HeaderLastJP(), HeaderVwapJP(), HeaderPrevCloseJP(), HeaderPreopenBidJP(), HeaderPreopenAskJP(), HeaderPreopenMidJP(), HeaderLiveGapBpJP(), HeaderLiveGapBucketJP(), HeaderLiveGapActionJP())
End Function

Private Function CandidateHeaderList() As Variant
    CandidateHeaderList = Array("TickerSrc", "Selected", "SignalMode", "Session", "ATR_n", "TPk", "SLk", "J_th", "ForwardPF", "ForwardTrades", "ForwardWin", "MaxDD", "WinCI_L", "WinCI_H", "ExpBootMean", "ExpBootLow", "ExpBootHigh", "ForwardAvgBars", "GapBucket", "GapRule", "GapSummary", "DynamicQty", "PlanTag", "EntryBuyPx", "EntrySellPx", "TP_price", "SL_price", "EntryStatus", "EntrySide")
End Function

Private Function DashboardHeaderList() As Variant
    Dim rt As Variant: rt = RealTimeHeaderList()
    Dim cand As Variant: cand = CandidateHeaderList()
    Dim total() As Variant
    Dim i As Long, offset As Long
    ReDim total(0 To UBound(rt) + UBound(cand) + 1)
    For i = 0 To UBound(rt)
        total(offset) = rt(i)
        offset = offset + 1
    Next i
    Dim j As Long
    For j = 0 To UBound(cand)
        total(offset) = cand(j)
        offset = offset + 1
    Next j
    DashboardHeaderList = total
End Function


Private Function HeaderNameJP() As String
    HeaderNameJP = ChrW(&H9298) & ChrW(&H67C4) & ChrW(&H540D)
End Function

Private Function HeaderLastJP() As String
    HeaderLastJP = ChrW(&H73FE) & ChrW(&H5728) & ChrW(&H5024)
End Function

Private Function HeaderJValueJP() As String
    HeaderJValueJP = ChrW(&H73FE) & ChrW(&H5728) & ChrW(&H306E) & "J" & ChrW(&H5024)
End Function

Private Function HeaderJGapJP() As String
    HeaderJGapJP = ChrW(&H95BE) & ChrW(&H5024) & ChrW(&H4E56) & ChrW(&H96E2) & ChrW(&H7387) & ChrW(&HFF08) & "%" & ChrW(&HFF09)
End Function

Private Function HeaderVwapJP() As String
    HeaderVwapJP = ChrW(&H51FA) & ChrW(&H6765) & ChrW(&H9AD8) & ChrW(&H52A0) & ChrW(&H91CD) & ChrW(&H5E73) & ChrW(&H5747)
End Function

Private Function HeaderSignalStatusJP() As String
    HeaderSignalStatusJP = ChrW(&H30B7) & ChrW(&H30B0) & ChrW(&H30CA) & ChrW(&H30EB) & ChrW(&H70B9) & ChrW(&H706F)
End Function

Private Function HeaderSignalKindJP() As String
    HeaderSignalKindJP = ChrW(&H30B7) & ChrW(&H30B0) & ChrW(&H30CA) & ChrW(&H30EB) & ChrW(&H7A2E) & ChrW(&H5225)
End Function

Private Function HeaderPrevCloseJP() As String
    HeaderPrevCloseJP = ChrW(&H524D) & ChrW(&H65E5) & ChrW(&H7D42) & ChrW(&H5024)
End Function

Private Function HeaderPreopenBidJP() As String
    HeaderPreopenBidJP = ChrW(&H6700) & ChrW(&H826F) & ChrW(&H8CB7) & ChrW(&H6C17) & ChrW(&H914D) & ChrW(&H5024)
End Function

Private Function HeaderPreopenAskJP() As String
    HeaderPreopenAskJP = ChrW(&H6700) & ChrW(&H826F) & ChrW(&H58F2) & ChrW(&H6C17) & ChrW(&H914D) & ChrW(&H5024)
End Function

Private Function HeaderPreopenMidJP() As String
    HeaderPreopenMidJP = ChrW(&H6C17) & ChrW(&H914D) & ChrW(&H5024) & ChrW(&HFF08) & ChrW(&H4E2D) & ChrW(&H592E) & ChrW(&HFF09)
End Function

Private Function RssFieldPreopenBidCode() As String
    RssFieldPreopenBidCode = "56"
End Function

Private Function RssFieldPreopenAskCode() As String
    RssFieldPreopenAskCode = "55"
End Function

Private Function HeaderLiveGapBpJP() As String
    HeaderLiveGapBpJP = ChrW(&H30E9) & ChrW(&H30A4) & ChrW(&H30D6) & ChrW(&H30AE) & ChrW(&H30E3) & ChrW(&H30C3) & ChrW(&H30D7) & "(bp)"
End Function

Private Function HeaderLiveGapBucketJP() As String
    HeaderLiveGapBucketJP = ChrW(&H30E9) & ChrW(&H30A4) & ChrW(&H30D6) & ChrW(&H30AE) & ChrW(&H30E3) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H5E2F)
End Function

Private Function HeaderLiveGapActionJP() As String
    HeaderLiveGapActionJP = ChrW(&H30E9) & ChrW(&H30A4) & ChrW(&H30D6) & ChrW(&H30A2) & ChrW(&H30AF) & ChrW(&H30B7) & ChrW(&H30E7) & ChrW(&H30F3)
End Function

Private Function RssFieldNameFullJP() As String
    RssFieldNameFullJP = ChrW(&H9298) & ChrW(&H67C4) & ChrW(&H540D) & ChrW(&H79F0)
End Function

Private Function RssFieldVwapFullJP() As String
    RssFieldVwapFullJP = ChrW(&H51FA) & ChrW(&H6765) & ChrW(&H9AD8) & ChrW(&H52A0) & ChrW(&H91CD) & ChrW(&H5E73) & ChrW(&H5747)
End Function

Private Function RssFieldLastJP() As String
    RssFieldLastJP = HeaderLastJP()
End Function

Private Function RssFieldBestAskJP() As String
    RssFieldBestAskJP = ChrW(&H6700) & ChrW(&H826F) & ChrW(&H58F2) & ChrW(&H6C17) & ChrW(&H914D) & ChrW(&H5024)
End Function

Private Function RssFieldBestBidJP() As String
    RssFieldBestBidJP = ChrW(&H6700) & ChrW(&H826F) & ChrW(&H8CB7) & ChrW(&H6C17) & ChrW(&H914D) & ChrW(&H5024)
End Function


Public Sub ButtonLoadCandidates()
    On Error GoTo Fail
    EnsureRuntimeReady False
    LoadCandidates
    Exit Sub
Fail:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    LogDebug "ButtonLoadCandidates error err=" & errNum & " desc=" & errDesc
    Err.Clear
    If errNum <> 0 Then Err.Raise errNum, "AutoTrader.ButtonLoadCandidates", errDesc
End Sub

Private Sub ApplyRealtimeColumns(ByVal ws As Worksheet, Optional ByVal fillLastOverride As Long = 0)

    Dim wasProtected As Boolean
    On Error Resume Next
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect Password:=""
    On Error GoTo 0


    Dim tickerCol As Long

    tickerCol = FindColumn(ws, DASH_HEADER_ROW, "Ticker")

    If tickerCol = 0 Then Exit Sub



    Dim lastDataRow As Long

    lastDataRow = ws.Cells(ws.Rows.Count, tickerCol).End(xlUp).ROW

    If lastDataRow < DASH_DATA_START Then lastDataRow = DASH_DATA_START



    Dim fillLast As Long

    fillLast = lastDataRow

    If fillLast < DASH_DATA_START + DASH_FORMULA_ROWS Then

        fillLast = DASH_DATA_START + DASH_FORMULA_ROWS

    End If

    If fillLastOverride >= DASH_DATA_START Then

        fillLast = fillLastOverride

    End If



    Dim tickerSrcCol As Long
    tickerSrcCol = FindColumn(ws, DASH_HEADER_ROW, "TickerSrc")
    If tickerSrcCol > 0 Then
        Dim tickerRef As String
        tickerRef = BuildR1C1Ref(tickerSrcCol, tickerCol)
        SetColumnFormula ws, tickerCol, fillLast, "=IF(" & tickerRef & "="""",""""," & tickerRef & ")"
    Else
        ws.Range(ws.Cells(DASH_DATA_START, tickerCol), ws.Cells(fillLast, tickerCol)).ClearContents
    End If



    Dim signalStatusCol As Long

    signalStatusCol = FindColumn(ws, DASH_HEADER_ROW, HeaderSignalStatusJP())
    If signalStatusCol = 0 And tickerCol > 0 Then signalStatusCol = tickerCol + 4

    Dim signalKindCol As Long

    signalKindCol = FindColumn(ws, DASH_HEADER_ROW, HeaderSignalKindJP())
    If signalKindCol = 0 And tickerCol > 0 Then signalKindCol = tickerCol + 5

    Dim selectedCol As Long
    selectedCol = FindColumn(ws, DASH_HEADER_ROW, "Selected")

    Dim signalModeCol As Long
    signalModeCol = FindColumn(ws, DASH_HEADER_ROW, "SignalMode")

    Dim currentJCol As Long

    currentJCol = FindColumn(ws, DASH_HEADER_ROW, HeaderJValueJP())
    If currentJCol = 0 And tickerCol > 0 Then currentJCol = tickerCol + 2



    Dim gapPctCol As Long

    gapPctCol = FindColumn(ws, DASH_HEADER_ROW, HeaderJGapJP())
    If gapPctCol = 0 And tickerCol > 0 Then gapPctCol = tickerCol + 3



    Dim nameCol As Long

    nameCol = FindColumn(ws, DASH_HEADER_ROW, HeaderNameJP())
    If nameCol = 0 And tickerCol > 0 Then nameCol = tickerCol + 1

    If nameCol > 0 Then

        Dim nameRef As String

        nameRef = BuildR1C1Ref(tickerCol, nameCol)

        SetColumnFormula ws, nameCol, fillLast, "=IF(" & nameRef & "="","",IFERROR(RssMarket(" & nameRef & "," & QuoteForFormula(RssFieldNameFullJP()) & "),""))"

    End If



    Dim lastCol As Long

    lastCol = FindColumn(ws, DASH_HEADER_ROW, HeaderLastJP())
    If lastCol = 0 And tickerCol > 0 Then lastCol = tickerCol + 6

    If lastCol > 0 Then

        Dim lastRef As String

        lastRef = BuildR1C1Ref(tickerCol, lastCol)

        SetColumnFormula ws, lastCol, fillLast, "=IF(" & lastRef & "="","",IFERROR(RssMarket(" & lastRef & "," & QuoteForFormula(HeaderLastJP()) & "),""))"

    End If



    Dim vwapCol As Long

    vwapCol = FindColumn(ws, DASH_HEADER_ROW, HeaderVwapJP())
    If vwapCol = 0 And tickerCol > 0 Then vwapCol = tickerCol + 7

    If vwapCol > 0 Then

        Dim vwapRef As String

        vwapRef = BuildR1C1Ref(tickerCol, vwapCol)

        SetColumnFormula ws, vwapCol, fillLast, "=IF(" & vwapRef & "="","",IFERROR(RssMarket(" & vwapRef & "," & QuoteForFormula(RssFieldVwapFullJP()) & "),""))"

    End If



    Dim atrCol As Long

    atrCol = FindColumn(ws, DASH_HEADER_ROW, "ATR_n")

    If currentJCol > 0 And lastCol > 0 And vwapCol > 0 And atrColEntry > 0 Then

        Dim lastRefJ As String

        Dim vwapRefJ As String

        Dim atrRefJ As String

        lastRefJ = BuildR1C1Ref(lastCol, currentJCol)

        vwapRefJ = BuildR1C1Ref(vwapCol, currentJCol)

        atrRefJ = BuildR1C1Ref(atrCol, currentJCol)

        SetColumnFormula ws, currentJCol, fillLast, "=IF(OR(" & lastRefJ & "=""," & vwapRefJ & "=""," & atrRefJ & "=0),"",((" & lastRefJ & "-" & vwapRefJ & ")/" & atrRefJ & ")/100)"

    ElseIf currentJCol > 0 Then

        SetColumnFormula ws, currentJCol, fillLast, "="""

    End If



    Dim prevCloseCol As Long

    prevCloseCol = FindColumn(ws, DASH_HEADER_ROW, HeaderPrevCloseJP())
    If prevCloseCol = 0 And tickerCol > 0 Then prevCloseCol = tickerCol + 8

    If prevCloseCol > 0 Then

        Dim prevRef As String

        prevRef = BuildR1C1Ref(tickerCol, prevCloseCol)

        SetColumnFormula ws, prevCloseCol, fillLast, "=IF(" & prevRef & "="","",IFERROR(RssMarket(" & prevRef & "," & QuoteForFormula(HeaderPrevCloseJP()) & "),""))"

    End If



    Dim bidCol As Long

    bidCol = FindColumn(ws, DASH_HEADER_ROW, HeaderPreopenBidJP())
    If bidCol = 0 And tickerCol > 0 Then bidCol = tickerCol + 9

    If bidCol > 0 Then

        Dim bidRef As String

        bidRef = BuildR1C1Ref(tickerCol, bidCol)

        ' Bid/Ask 邵ｺ・ｯ鬩ｫ菫ｶ豌帷ｹｧ・ｳ郢晢ｽｼ郢晉判譫夊氛諤懊・郢ｧ蛛ｵ笳守ｸｺ・ｮ邵ｺ・ｾ邵ｺ・ｾ雋ゑｽ｡邵ｺ蜻ｻ・ｼ繝ｻ85A.T 驕ｲ蟲ｨ繝ｻ郢ｧ・｢郢晢ｽｫ郢晁ｼ斐＜郢晏生繝｣郢晏現・定惺・ｫ郢ｧﾂ郢ｧ・ｳ郢晢ｽｼ郢晉甥・ｯ・ｾ陟｢諛ｶ・ｼ繝ｻ        SetColumnFormula ws, bidCol, fillLast, "=IF(" & bidRef & "="","",IFERROR(RssMarket(" & bidRef & "," & QuoteForFormula(RssFieldBestBidJP()) & "),""))"

    End If



    Dim askCol As Long

    askCol = FindColumn(ws, DASH_HEADER_ROW, HeaderPreopenAskJP())
    If askCol = 0 And tickerCol > 0 Then askCol = tickerCol + 10

    If askCol > 0 Then

        Dim askRef As String

        askRef = BuildR1C1Ref(tickerCol, askCol)

        SetColumnFormula ws, askCol, fillLast, "=IF(" & askRef & "="","",IFERROR(RssMarket(" & askRef & "," & QuoteForFormula(RssFieldBestAskJP()) & "),""))"

    End If



    Dim midCol As Long

    midCol = FindColumn(ws, DASH_HEADER_ROW, HeaderPreopenMidJP())
    If midCol = 0 And tickerCol > 0 Then midCol = tickerCol + 11

    If midCol > 0 And bidCol > 0 And askCol > 0 Then

        Dim bidRefForMid As String

        Dim askRefForMid As String

        bidRefForMid = BuildR1C1Ref(bidCol, midCol)

        askRefForMid = BuildR1C1Ref(askCol, midCol)

        SetColumnFormula ws, midCol, fillLast, "=IF(OR(" & bidRefForMid & "=""," & askRefForMid & "=""),"",(" & bidRefForMid & "+" & askRefForMid & ")/2)"

    End If



    Dim gapBpCol As Long

    gapBpCol = FindColumn(ws, DASH_HEADER_ROW, HeaderLiveGapBpJP())
    If gapBpCol = 0 And tickerCol > 0 Then gapBpCol = tickerCol + 12

    If gapBpCol > 0 And midCol > 0 And prevCloseCol > 0 Then

        Dim midRefForGap As String

        Dim prevRefForGap As String

        midRefForGap = BuildR1C1Ref(midCol, gapBpCol)

        prevRefForGap = BuildR1C1Ref(prevCloseCol, gapBpCol)

        SetColumnFormula ws, gapBpCol, fillLast, "=IF(OR(" & midRefForGap & "=""," & prevRefForGap & "=""),"",(" & midRefForGap & "-" & prevRefForGap & ")/" & prevRefForGap & "*10000)"

    End If



    Dim jthCol As Long

    jthCol = FindColumn(ws, DASH_HEADER_ROW, "J_th")

    If gapPctCol > 0 Then

        If currentJCol > 0 And jthCol > 0 Then

            Dim currentRefGap As String
            Dim jthRef As String
            currentRefGap = BuildR1C1Ref(currentJCol, gapPctCol)
            jthRef = BuildR1C1Ref(jthCol, gapPctCol)

            ' 鬮｣蛹・ｽｽ・ｵ髯樊ｻ薙・繝ｻ・ｱ繝ｻ・ｬ鬯ｩ髦ｪ繝ｻ郢晢ｽｻ郢晢ｽｻ繝ｻ・ｮ髯樊ｺ假ｽ代・・ｽ繝ｻ・ｾ郢晢ｽｻ繝ｻ・ｩ: |J - J_th| / |J_th| * 100
            SetColumnFormula ws, gapPctCol, fillLast, "=IF(OR(" & jthRef & "=""," & currentRefGap & "=""," & jthRef & "=0),"",ABS(" & currentRefGap & "-" & jthRef & ")/ABS(" & jthRef & ")*100)"

            Dim rngGap As Range
            Set rngGap = ws.Range(ws.Cells(DASH_DATA_START, gapPctCol), ws.Cells(fillLast, gapPctCol))

            On Error Resume Next
            rngGap.FormatConditions.Delete
            On Error GoTo 0

            ' 鬮ｫ・ｴ陞滂ｽｲ繝ｻ・ｽ繝ｻ・｡鬮｣豈費ｽｼ螟ｲ・ｽ・ｽ繝ｻ・ｶ鬮｣逧ｮ逕･・つ繝ｻ・･驕ｯ・ｶ繝ｻ・ｳ鬮ｫ・ｴ陷ｴ繝ｻ・ｽ・ｽ繝ｻ・ｸ鬮ｯ貊会ｽｻ・｣郢晢ｽｻ 0% 鬩搾ｽｵ繝ｻ・ｺ郢晢ｽｻ繝ｻ・ｯ鬮ｮ雜｣・ｽ・ｼ驛｢譎｢・ｽ・ｻ郢晢ｽｻ隶難ｽ｣邵ｺ讌｢閼ゅ・・｣繝ｻ縺､ﾂ驛｢譎｢・ｽ・ｻ0%鬮｣豈費ｽｼ螟ｲ・ｽ・ｽ繝ｻ・･鬮｣蛹・ｽｽ・ｳ髣包ｽｵ隴擾ｽｴ郢晢ｽｻ鬯ｮ・ｦ繝ｻ・ｮ驛｢譎｢・ｽ・ｻ郢晢ｽｻ隶難ｽ｣邵ｺ骰具ｽｹ譎｢・ｽ・ｻ
            Dim fc0 As FormatCondition
            Set fc0 = rngGap.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
            fc0.Interior.Color = RGB(0, 176, 80)
            On Error Resume Next: fc0.StopIfTrue = True: On Error GoTo 0

            Dim fc1 As FormatCondition
            Set fc1 = rngGap.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="30")
            fc1.Interior.Color = RGB(146, 208, 80)

        Else
            SetColumnFormula ws, gapPctCol, fillLast, "="""
        End If

    End If



    Dim gapBucketCol As Long

    gapBucketCol = FindColumn(ws, DASH_HEADER_ROW, HeaderLiveGapBucketJP())
    If gapBucketCol = 0 And tickerCol > 0 Then gapBucketCol = tickerCol + 13

    If gapBucketCol > 0 And gapBpCol > 0 Then

        Dim gapRef As String

        gapRef = BuildR1C1Ref(gapBpCol, gapBucketCol)

        SetColumnFormula ws, gapBucketCol, fillLast, "=IF(" & gapRef & "="""","""",IF(ABS(" & gapRef & ")>=120," & QuoteForFormula(">=120bp") & ",IF(ABS(" & gapRef & ")>=80," & QuoteForFormula("80-120bp") & ",IF(ABS(" & gapRef & ")>=50," & QuoteForFormula("50-80bp") & "," & QuoteForFormula("<50bp") & "))))"

    End If



    Dim actionCol As Long

    actionCol = FindColumn(ws, DASH_HEADER_ROW, HeaderLiveGapActionJP())
    If actionCol = 0 And tickerCol > 0 Then actionCol = tickerCol + 14

    If actionCol > 0 And gapBucketCol > 0 Then

        Dim bucketRef As String

        bucketRef = BuildR1C1Ref(gapBucketCol, actionCol)

        SetColumnFormula ws, actionCol, fillLast, "=IF(" & bucketRef & "="""","""",IF(" & bucketRef & "=" & QuoteForFormula(">=120bp") & "," & QuoteForFormula("j-cross only; TP-0.2; SL+0.2") & ",IF(" & bucketRef & "=" & QuoteForFormula("80-120bp") & "," & QuoteForFormula("Skip opposite; J_th+0.2") & ",IF(" & bucketRef & "=" & QuoteForFormula("50-80bp") & "," & QuoteForFormula("J_th+0.1") & "," & QuoteForFormula("Baseline") & "))))"

    End If

    If signalStatusCol > 0 And selectedCol > 0 And currentJCol > 0 And jthCol > 0 Then
        Dim selRef As String
        Dim jRef As String
        Dim jthStatusRef As String
        Dim bidRefStatus As String
        Dim askRefStatus As String
        Dim midRefStatus As String
        Dim prevRefStatus As String
        selRef = BuildR1C1Ref(selectedCol, signalStatusCol)
        jRef = BuildR1C1Ref(currentJCol, signalStatusCol)
        jthStatusRef = BuildR1C1Ref(jthCol, signalStatusCol)
        bidRefStatus = "0"
        If bidCol > 0 Then bidRefStatus = BuildR1C1Ref(bidCol, signalStatusCol)
        askRefStatus = "0"
        If askCol > 0 Then askRefStatus = BuildR1C1Ref(askCol, signalStatusCol)
        midRefStatus = "0"
        If midCol > 0 Then midRefStatus = BuildR1C1Ref(midCol, signalStatusCol)
        prevRefStatus = "0"
        If prevCloseCol > 0 Then prevRefStatus = BuildR1C1Ref(prevCloseCol, signalStatusCol)
        Dim statusFormula As String
        statusFormula = _
            "=IF(" & selRef & "<>1,""""," & _
                "IF(OR(" & jRef & "=""""," & jthStatusRef & "=""""," & jthStatusRef & "=0),""""," & _
                    "IF(" & jRef & "<0," & _
                        "IF(IF(" & askRefStatus & ">0," & askRefStatus & ",IF(" & midRefStatus & ">0," & midRefStatus & "," & prevRefStatus & "))<=0,""NO_PRICE"",IF(ABS(" & jRef & ")>=ABS(" & jthStatusRef & "),""BUY"",""""))," & _
                        "IF(IF(" & bidRefStatus & ">0," & bidRefStatus & ",IF(" & midRefStatus & ">0," & midRefStatus & "," & prevRefStatus & "))<=0,""NO_PRICE"",IF(ABS(" & jRef & ")>=ABS(" & jthStatusRef & "),""SELL"","""")))))"
        SetColumnFormula ws, signalStatusCol, fillLast, statusFormula
    End If

    If signalKindCol > 0 And signalStatusCol > 0 And signalModeCol > 0 Then
        Dim statusRef As String
        Dim modeRef As String

        statusRef = BuildR1C1Ref(signalStatusCol, signalKindCol)
        modeRef = BuildR1C1Ref(signalModeCol, signalKindCol)

        Dim kindFormula As String
        kindFormula = _
            "=IF(" & statusRef & "="""",""""," & _
                "IF(" & statusRef & "=""NO_PRICE"",""NO_PRICE"",IF(OR(" & statusRef & "=""BUY""," & statusRef & "=""SELL""),IF(" & modeRef & "=""""," & statusRef & "," & statusRef & " & "" / "" & " & modeRef & ")," & statusRef & "))))"
        SetColumnFormula ws, signalKindCol, fillLast, kindFormula
    End If

    ' Entry/TP/SL 髯ｦ・ｨ驕会ｽｺ陋ｻ蜉ｱ・帝坎閧ｲ・ｮ繝ｻ    Dim entryBuyCol As Long: entryBuyCol = FindColumn(ws, DASH_HEADER_ROW, "EntryBuyPx")
    Dim entrySellCol As Long: entrySellCol = FindColumn(ws, DASH_HEADER_ROW, "EntrySellPx")
    Dim tpPriceCol As Long: tpPriceCol = FindColumn(ws, DASH_HEADER_ROW, "TP_price")
    Dim slPriceCol As Long: slPriceCol = FindColumn(ws, DASH_HEADER_ROW, "SL_price")
    Dim entrySideCol As Long: entrySideCol = FindColumn(ws, DASH_HEADER_ROW, "EntrySide")
    Dim atrColEntry As Long: atrColEntry = FindColumn(ws, DASH_HEADER_ROW, "ATR")
    Dim tpKColEntry As Long: tpKColEntry = FindColumn(ws, DASH_HEADER_ROW, "TPk")
    Dim slKColEntry As Long: slKColEntry = FindColumn(ws, DASH_HEADER_ROW, "SLk")
    Dim jValCol As Long: jValCol = FindColumn(ws, DASH_HEADER_ROW, HeaderJValueJP())

    If entryBuyCol > 0 And vwapCol > 0 And jthCol > 0 And atrColEntry > 0 Then
        Dim vwapRefEB As String: vwapRefEB = BuildR1C1Ref(vwapCol, entryBuyCol)
        Dim jthRefEB As String: jthRefEB = BuildR1C1Ref(jthCol, entryBuyCol)
        Dim atrRefEB As String: atrRefEB = BuildR1C1Ref(atrColEntry, entryBuyCol)
        SetColumnFormula ws, entryBuyCol, fillLast, "=IF(OR(" & vwapRefEB & "="""," & jthRefEB & "="""," & atrRefEB & "=0),""", " & vwapRefEB & "-ABS(" & jthRefEB & ")*" & atrRefEB & ")"
    End If
    If entrySellCol > 0 And vwapCol > 0 And jthCol > 0 And atrColEntry > 0 Then
        Dim vwapRefES As String: vwapRefES = BuildR1C1Ref(vwapCol, entrySellCol)
        Dim jthRefES As String: jthRefES = BuildR1C1Ref(jthCol, entrySellCol)
        Dim atrRefES As String: atrRefES = BuildR1C1Ref(atrColEntry, entrySellCol)
        SetColumnFormula ws, entrySellCol, fillLast, "=IF(OR(" & vwapRefES & "="""," & jthRefES & "="""," & atrRefES & "=0),""", " & vwapRefES & "+ABS(" & jthRefES & ")*" & atrRefES & ")"
    End If
    If entrySideCol > 0 And jValCol > 0 Then
        Dim jRefESide As String: jRefESide = BuildR1C1Ref(jValCol, entrySideCol)
        SetColumnFormula ws, entrySideCol, fillLast, "=IF(" & jRefESide & "=""",""",IF(" & jRefESide & "<0,\"BUY\",\"SELL\"))"
    End If
    If tpPriceCol > 0 And slPriceCol > 0 And atrColEntry > 0 And tpKCol > 0 And slKCol > 0 And jValCol > 0 And entryBuyCol > 0 And entrySellCol > 0 Then
        Dim jRefTP As String: jRefTP = BuildR1C1Ref(jValCol, tpPriceCol)
        Dim atrRefTP As String: atrRefTP = BuildR1C1Ref(atrColEntry, tpPriceCol)
        Dim tpKRef As String: tpKRef = BuildR1C1Ref(tpKColEntry, tpPriceCol)
        Dim eBuyRef As String: eBuyRef = BuildR1C1Ref(entryBuyCol, tpPriceCol)
        Dim eSellRef As String: eSellRef = BuildR1C1Ref(entrySellCol, tpPriceCol)
        SetColumnFormula ws, tpPriceCol, fillLast, "=IF(OR(" & atrRefTP & "=0," & tpKRef & "="""),""",IF(" & jRefTP & "<0," & eBuyRef & "+" & tpKRef & "*" & atrRefTP & "," & eSellRef & "-" & tpKRef & "*" & atrRefTP & "))"

        Dim jRefSL As String: jRefSL = BuildR1C1Ref(jValCol, slPriceCol)
        Dim atrRefSL As String: atrRefSL = BuildR1C1Ref(atrColEntry, slPriceCol)
        Dim slKRef As String: slKRef = BuildR1C1Ref(slKColEntry, slPriceCol)
        Dim eBuyRef2 As String: eBuyRef2 = BuildR1C1Ref(entryBuyCol, slPriceCol)
        Dim eSellRef2 As String: eSellRef2 = BuildR1C1Ref(entrySellCol, slPriceCol)
        SetColumnFormula ws, slPriceCol, fillLast, "=IF(OR(" & atrRefSL & "=0," & slKRef & "="""),""",IF(" & jRefSL & "<0," & eBuyRef2 & "-" & slKRef & "*" & atrRefSL & "," & eSellRef2 & "+" & slKRef & "*" & atrRefSL & "))"
    End If


    If wasProtected Then
        ProtectRealtimeColumns ws
    End If

End Sub



Private Sub EnsureRealtimeFirstRow(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    Dim tickerCol As Long
    tickerCol = FindColumn(ws, DASH_HEADER_ROW, "Ticker")
    If tickerCol = 0 Then Exit Sub

    Dim hasFormula As Boolean
    On Error Resume Next
    hasFormula = ws.Cells(DASH_DATA_START, tickerCol).HasFormula
    On Error GoTo 0
    If hasFormula Then Exit Sub

    Dim wasProtected As Boolean
    On Error Resume Next
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect Password:=""
    On Error GoTo 0

    ApplyRealtimeColumns ws

    If wasProtected Then
        ProtectRealtimeColumns ws
    End If

End Sub



Private Sub SetColumnFormula(ByVal ws As Worksheet, ByVal col As Long, ByVal fillLast As Long, ByVal formulaR1C1 As String)

    If col <= 0 Then Exit Sub

    Dim rngTarget As Range
    Set rngTarget = ws.Range(ws.Cells(DASH_DATA_START, col), ws.Cells(fillLast, col))

    Dim firstCell As Range
    Set firstCell = rngTarget.Cells(1, 1)

    Dim hasExisting As Boolean
    Dim existingFormula As String
    On Error Resume Next
    hasExisting = firstCell.HasFormula
    If hasExisting Then existingFormula = firstCell.FormulaR1C1
    On Error GoTo 0

    If hasExisting Then
        If StrComp(existingFormula, formulaR1C1, vbBinaryCompare) = 0 Then
            Exit Sub
        End If
    End If

    On Error GoTo UseLocal

    firstCell.FormulaR1C1 = formulaR1C1
    GoTo AfterApply

UseLocal:
    Err.Clear
    On Error GoTo ApplyFailed
    firstCell.FormulaR1C1Local = formulaR1C1

AfterApply:
    If rngTarget.Rows.Count > 1 Then
        firstCell.AutoFill Destination:=rngTarget
    End If

    If Not firstCell.HasFormula Then
        LogDebug "SetColumnFormula missing formula col=" & col & " after apply formulaR1C1=" & formulaR1C1
    End If

    On Error GoTo 0
    Exit Sub

ApplyFailed:
    LogDebug "SetColumnFormula error col=" & col & " err=" & Err.Number & " desc=" & Err.Description & " formulaR1C1=" & formulaR1C1
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub LogDebug(ByVal message As String)
    Dim folder As String
    folder = ThisWorkbook.path & "\logs"
    On Error Resume Next
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    Dim path As String
    path = folder & "\autotrader_debug.log"

    Dim f As Integer
    f = FreeFile

    On Error Resume Next
    Open path For Append As #f
    If Err.Number = 0 Then
        Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss"), message
        Close #f
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function IsHeadless() As Boolean

    On Error Resume Next

    IsHeadless = (Application.Interactive = False)

    On Error GoTo 0

End Function




Private Function BuildR1C1Ref(ByVal sourceCol As Long, ByVal targetCol As Long) As String

    Dim offset As Long

    offset = sourceCol - targetCol

    If offset = 0 Then

        BuildR1C1Ref = "RC"

    Else

        BuildR1C1Ref = "RC[" & CStr(offset) & "]"

    End If

End Function



Private Function ResolveHeaderAliases(ByVal name As String) As Variant

    Dim key As String

    key = LCase$(name)



    Dim headerNameKey As String

    headerNameKey = LCase$(HeaderNameJP())

    Dim rssNameKey As String

    rssNameKey = LCase$(RssFieldNameFullJP())

    Dim headerLastKey As String

    headerLastKey = LCase$(HeaderLastJP())

    Dim headerVwapKey As String

    headerVwapKey = LCase$(HeaderVwapJP())

    Dim rssVwapKey As String

    rssVwapKey = LCase$(RssFieldVwapFullJP())



    Select Case key

        Case "ticker"

            ResolveHeaderAliases = Array("Ticker")

        Case "j", LCase$(HeaderJValueJP()), "current_j"

            ResolveHeaderAliases = Array(HeaderJValueJP(), "J", "current_j")

        Case headerNameKey, rssNameKey, "name"

            ResolveHeaderAliases = Array(HeaderNameJP(), RssFieldNameFullJP(), "Name")

        Case LCase$(HeaderJGapJP()), "jgap", "gap_pct"

            ResolveHeaderAliases = Array(HeaderJGapJP(), "JGap", "gap_pct")

        Case LCase$(HeaderSignalStatusJP()), "signalstatus"

            ResolveHeaderAliases = Array(HeaderSignalStatusJP(), "SignalStatus")

        Case LCase$(HeaderSignalKindJP()), "signalkind"

            ResolveHeaderAliases = Array(HeaderSignalKindJP(), "SignalKind")

        Case headerLastKey, "last"

            ResolveHeaderAliases = Array(HeaderLastJP(), RssFieldLastJP(), "Last")

        Case headerVwapKey, rssVwapKey, "vwap"

            ResolveHeaderAliases = Array(HeaderVwapJP(), RssFieldVwapFullJP(), "VWAP")

        Case "forwardpf", "forward_pf_eff"

            ResolveHeaderAliases = Array("ForwardPF", "forward_pf_eff")

        Case "forwardtrades", "forward_trades"

            ResolveHeaderAliases = Array("ForwardTrades", "forward_trades")

        Case "forwardwin", "forward_winrate"

            ResolveHeaderAliases = Array("ForwardWin", "forward_winrate")

        Case "prevclose"

            ResolveHeaderAliases = Array("PrevClose", HeaderPrevCloseJP())

        Case "preopenbid"

            ResolveHeaderAliases = Array("PreOpenBid", HeaderPreopenBidJP())

        Case "preopenask"

            ResolveHeaderAliases = Array("PreOpenAsk", HeaderPreopenAskJP())

        Case "preopenmid"

            ResolveHeaderAliases = Array("PreOpenMid", HeaderPreopenMidJP())

        Case "livegapbp"

            ResolveHeaderAliases = Array("LiveGapBp", HeaderLiveGapBpJP())

        Case "livegapbucket"

            ResolveHeaderAliases = Array("LiveGapBucket", HeaderLiveGapBucketJP())

        Case "livegapaction"

            ResolveHeaderAliases = Array("LiveGapAction", HeaderLiveGapActionJP())

        Case Else

            ResolveHeaderAliases = Array(name)

    End Select

End Function






Public Sub ButtonPushCandidates()
    EnsureRuntimeReady True
    PushCandidatesToDashboard
End Sub

Public Sub ButtonStartAuto()
    EnsureRuntimeReady True
    StartAutoTrading
End Sub

Public Sub ButtonStopAuto()
    StopAutoTrading
End Sub

Public Sub ResetDashboardHeaders()
    ' Rebuild header labels only; keep existing formulas intact.
    setupInitialized = False
    EnsureSetup
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    WriteHeaderTexts ws
End Sub

Private Sub ClearDashboardRealtime(ByVal ws As Worksheet)
    Dim baseCol As Long: baseCol = 8
    Dim rtHeaders As Variant: rtHeaders = RealTimeHeaderList()
    Dim lastCol As Long: lastCol = baseCol + UBound(rtHeaders)
    ws.Range(ws.Cells(DASH_DATA_START, baseCol), ws.Cells(DASH_DATA_START + DASH_FORMULA_ROWS, lastCol)).ClearContents
End Sub

Public Sub ButtonRefreshNow()
    EnsureRuntimeReady True
    EvaluateAndQueueOrders
End Sub

Public Sub ButtonQueueNow()
    On Error GoTo Fail
    EnsureRuntimeReady True
    QueueNowDryRun
    Exit Sub
Fail:
    LogDebug "ButtonQueueNow error err=" & Err.Number & " desc=" & Err.Description
    Err.Clear
End Sub

Private Sub QueueNowDryRun()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim tCol As Long: tCol = FindColumn(ws, DASH_HEADER_ROW, "Ticker")
    Dim selCol As Long: selCol = FindColumn(ws, DASH_HEADER_ROW, "Selected")
    Dim eBuyCol As Long: eBuyCol = FindColumn(ws, DASH_HEADER_ROW, "EntryBuyPx")
    Dim eSellCol As Long: eSellCol = FindColumn(ws, DASH_HEADER_ROW, "EntrySellPx")
    If tCol = 0 Or selCol = 0 Or eBuyCol = 0 Or eSellCol = 0 Then Exit Sub
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, tCol).End(xlUp).ROW
    For r = DASH_DATA_START To lastRow
        If CStr(ws.Cells(r, tCol).value) <> "" And CLng(IfZero(ws.Cells(r, selCol).value, 0)) = 1 Then
            Dim t As String: t = CStr(ws.Cells(r, tCol).value)
            Dim pxSell As Double: pxSell = CDbl(IfZero(ws.Cells(r, eSellCol).value, 0))
            Dim pxBuy As Double: pxBuy = CDbl(IfZero(ws.Cells(r, eBuyCol).value, 0))
            Dim qty As Long: qty = CLng(IfZero(ws.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))
            If pxSell > 0 Then PlaceOrderDryRun t, "SELL", pxSell, qty, "QUEUE_NOW:ENTRY_SELL"
            If pxBuy > 0 Then PlaceOrderDryRun t, "BUY", pxBuy, qty, "QUEUE_NOW:ENTRY_BUY"
            Exit For
        End If
    Next r
End Sub

Public Sub AttachFormulasFromDashboardTemplate()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "Template sheet 'Dashboard' not found. Skipping attach.", vbInformation
        Exit Sub
    End If
    Set wsDst = EnsureSheet(SHEET_DASHBOARD)
    Dim headers As Variant
    headers = Array(HeaderLastJP(), HeaderVwapJP(), "ATR")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        Dim h As String: h = CStr(headers(i))
        Dim cSrc As Long: cSrc = FindColumn(wsSrc, DASH_HEADER_ROW, h)
        Dim cDst As Long: cDst = FindColumn(wsDst, DASH_HEADER_ROW, h)
        If cSrc > 0 And cDst > 0 Then
            Dim f As String
            f = wsSrc.Cells(DASH_DATA_START, cSrc).Formula
            If Len(f) > 0 Then
                Dim lastRow As Long
                lastRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).ROW
                If lastRow >= DASH_DATA_START Then
                    wsDst.Range(wsDst.Cells(DASH_DATA_START, cDst), wsDst.Cells(lastRow, cDst)).Formula = f
                End If
            End If
        End If
    Next i
    wsDst.Columns.AutoFit
    MsgBox "Formulas attached from Dashboard template.", vbInformation
End Sub

Public Sub ButtonCatchUp()
    On Error Resume Next
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim root As String
    root = ThisWorkbook.path
    Dim cmd As String
    cmd = "cmd /c """ & root & "\scripts\run_nightly_now.cmd"""
    sh.Run cmd, 0, False
    MsgBox "Catch-up launched.", vbInformation
End Sub

Public Sub CancelNightlyStatusRefresh()
    On Error Resume Next
    NightlyStatus.StopNightlyMonitor
    On Error GoTo 0
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    EnsureSetup
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    ws.Range(DASH_STATUS_CELL).value = 0
End Sub

Private Sub EnsureSetup()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    If setupInitialized Then Exit Sub
    If wsDash.Range(DASH_MAX_ORDERS_CELL).value = "" Then
        wsDash.Range(DASH_MAX_ORDERS_CELL).value = DEFAULT_MAX_ORDERS
    End If
    If wsDash.Range(DASH_SESSION_START_CELL).value = "" Then
        wsDash.Range(DASH_SESSION_START_CELL).value = DEFAULT_SESSION_START
    End If
    If wsDash.Range(DASH_SESSION_END_CELL).value = "" Then
        wsDash.Range(DASH_SESSION_END_CELL).value = DEFAULT_SESSION_END
    End If
    If wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value = "" Then
        wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value = DEFAULT_SELECTED_DEFAULT
    End If
    If wsDash.Range(DASH_REENTRY_CELL).value = "" Then
        wsDash.Range(DASH_REENTRY_CELL).value = DEFAULT_REENTRY
    End If
    If wsDash.Range(DASH_HARDSTOP_CELL).value = "" Then
        wsDash.Range(DASH_HARDSTOP_CELL).value = DEFAULT_HARDSTOP
    End If
    If wsDash.Range(DASH_LIVE_CELL).value = "" Then
        wsDash.Range(DASH_LIVE_CELL).value = DEFAULT_LIVE
    End If
    If wsDash.Range(DASH_CLOSE_TIME_CELL).value = "" Then
        wsDash.Range(DASH_CLOSE_TIME_CELL).value = DEFAULT_CLOSE_TIME
    End If
    If wsDash.Range(DASH_QTY_CELL).value = "" Then
        wsDash.Range(DASH_QTY_CELL).value = DEFAULT_ORDER_QTY
    End If
    If wsDash.Range(DASH_BUDGET_CELL).value = "" Then
        wsDash.Range(DASH_BUDGET_CELL).value = DEFAULT_MAX_BUDGET
    End If
    If wsDash.Range(DASH_LOT_STEP_CELL).value = "" Then
        wsDash.Range(DASH_LOT_STEP_CELL).value = DEFAULT_LOT_STEP
    End If
    If wsDash.Range(DASH_SLIP_BP_CELL).value = "" Then
        wsDash.Range(DASH_SLIP_BP_CELL).value = DEFAULT_SLIP_BP
    End If
    ' 隴・ｽｰ髫募臆・ｨ・ｭ陞ｳ螢ｹ繝ｻ隴鯉ｽ｢陞ｳ螢ｼﾂ・､
    If wsDash.Range(DASH_TRACE_CELL).value = "" Then wsDash.Range(DASH_TRACE_CELL).value = 0
    If wsDash.Range(DASH_CANCEL_AT_END_CELL).value = "" Then wsDash.Range(DASH_CANCEL_AT_END_CELL).value = 0
    If wsDash.Range(DASH_ENTRY_GRACE_MIN_CELL).value = "" Then wsDash.Range(DASH_ENTRY_GRACE_MIN_CELL).value = 15
    If wsDash.Range(DASH_OCO_UPDATE_SEC_CELL).value = "" Then wsDash.Range(DASH_OCO_UPDATE_SEC_CELL).value = 30
    EnsureSheet SHEET_CANDIDATES
    EnsureOrdersSheet
    If Not DashboardHeadersReady(wsDash) Then
        EnsureHeaders wsDash
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateBeforeSave = True
    wsDash.Range("A2").value = "AutoTrade (0/1)"
    wsDash.Range("A3").value = "Daily Max Orders"
    wsDash.Range("A4").value = "Session Start"
    wsDash.Range("A5").value = "Session End"
    wsDash.Range("A6").value = "Selected Default (0/1)"
    wsDash.Range("A7").value = "Reentry Allowed (0/1)"
    wsDash.Range("A8").value = "Hard Stop (0/1)"
    wsDash.Range("A9").value = "Live Orders (0/1)"
    wsDash.Range("A10").value = "Order Macro Name"
    wsDash.Range("A11").value = "Close-Out Time (HH:MM:SS)"
    wsDash.Range("A12").value = "Order Quantity (Fallback)"
    wsDash.Range("A14").value = "Total Budget (JPY)"
    wsDash.Range("A15").value = "Lot Step Size (shares)"
    wsDash.Range("A16").value = "Max Slippage (bp)"
    wsDash.Range("A17:A40").ClearContents
    wsDash.Range("A17").value = "Trace (0/1)"
    wsDash.Range("A18").value = "Cancel at Session End (0/1)"
    wsDash.Range("A19").value = "Entry Grace Minutes"
    wsDash.Range("A20").value = "OCO Update Seconds"
    If sessionRiskUsed Is Nothing Then Set sessionRiskUsed = CreateObject("Scripting.Dictionary") Else sessionRiskUsed.RemoveAll
    If tickerRiskUsed Is Nothing Then Set tickerRiskUsed = CreateObject("Scripting.Dictionary") Else tickerRiskUsed.RemoveAll
    UpdateRiskUsageFromExecutions
    closeTimerScheduled = False
    closeTimer = 0
    setupInitialized = True
    EnsureQueueNowButton
End Sub

Public Sub EnsureQueueNowButton()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    On Error Resume Next
    Dim btn As Object
    For Each btn In ws.Buttons
        If btn.OnAction = "AutoTrader.ButtonQueueNow" Then Exit Sub
    Next
    On Error GoTo 0
    Dim target As Range
    Set target = ws.Range("D2")
    Dim newBtn As Object
    Set newBtn = ws.Buttons.Add(target.Left, target.Top, 140, 24)
    newBtn.Caption = "Queue Now (DryRun)"
    newBtn.OnAction = "AutoTrader.ButtonQueueNow"
End Sub
Private Sub EnsureRuntimeReady(Optional ByVal ensureDashHeaders As Boolean = False)
    If Not setupInitialized Then
        EnsureSetup
        Exit Sub
    End If
    EnsureSheet SHEET_CANDIDATES
    EnsureOrdersSheet
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    If ensureDashHeaders Then
        If Not DashboardHeadersReady(wsDash) Then
            EnsureHeaders wsDash
        End If
    End If
    EnsureRealtimeFirstRow wsDash
End Sub

Private Sub RebuildRiskUsageFromOrders()
    On Error GoTo Done
    If sessionRiskUsed Is Nothing Then Set sessionRiskUsed = CreateObject("Scripting.Dictionary") Else sessionRiskUsed.RemoveAll
    If tickerRiskUsed Is Nothing Then Set tickerRiskUsed = CreateObject("Scripting.Dictionary") Else tickerRiskUsed.RemoveAll
    totalRiskUsed = 0

    Dim wsDash As Worksheet: Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    Dim wsOrders As Worksheet: Set wsOrders = EnsureSheet(SHEET_ORDERS)
    Dim tCol As Long: tCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    Dim atrCol As Long: atrCol = FindColumn(wsDash, DASH_HEADER_ROW, "ATR_n")
    Dim slkCol As Long: slkCol = FindColumn(wsDash, DASH_HEADER_ROW, "SLk")
    Dim prevCol As Long: prevCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderPrevCloseJP())
    Dim midCol As Long: midCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderPreopenMidJP())
    If tCol = 0 Then GoTo Done

    Dim slipBp As Double
    slipBp = CDbl(IfZero(wsDash.Range(DASH_SLIP_BP_CELL).value, DEFAULT_SLIP_BP))

    Dim lastRow As Long
    lastRow = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).ROW
    Dim r As Long
    For r = 2 To lastRow
        Dim timeVal As Variant: timeVal = wsOrders.Cells(r, 1).value
        If IsDate(timeVal) Then
            If DateValue(timeVal) <> Date Then GoTo NextOrder
        End If
        Dim ticker As String: ticker = CStr(wsOrders.Cells(r, 2).value)
        Dim side As String: side = UCase$(CStr(wsOrders.Cells(r, 3).value))
        Dim price As Double: price = CDbl(IfZero(wsOrders.Cells(r, 4).value, 0))
        Dim qty As Long: qty = CLng(IfZero(wsOrders.Cells(r, 5).value, 0))
        Dim note As String: note = CStr(wsOrders.Cells(r, 6).value)
        If Len(ticker) = 0 Or qty <= 0 Then GoTo NextOrder
        Dim isClose As Boolean
        isClose = (InStr(1, note, "TP", vbTextCompare) > 0) Or (InStr(1, note, "SL", vbTextCompare) > 0) Or (InStr(1, note, "MOC", vbTextCompare) > 0) Or (InStr(1, note, "FLAT", vbTextCompare) > 0)

        Dim mode As String, sess As String
        mode = "": sess = ""
        If InStr(1, note, ":") > 0 Then
            mode = Split(note, ":")(0)
            If UBound(Split(note, ":")) >= 1 Then sess = Split(note, ":")(1)
        End If
        Dim sessionKey As String: sessionKey = GetSessionKey(sess, mode)

        ' Map ticker to dashboard row
        Dim dLast As Long: dLast = wsDash.Cells(wsDash.Rows.Count, tCol).End(xlUp).ROW
        Dim dr As Long, found As Long: found = 0
        For dr = DASH_DATA_START To dLast
            If CStr(wsDash.Cells(dr, tCol).value) = ticker Then found = dr: Exit For
        Next dr
        If found = 0 Then GoTo NextOrder

        Dim atr As Double: atr = CDbl(IfZero(wsDash.Cells(found, atrCol).value, 0))
        Dim slK As Double: slK = CDbl(IfZero(wsDash.Cells(found, slkCol).value, 0))
        Dim px As Double
        If midCol > 0 Then px = CDbl(IfZero(wsDash.Cells(found, midCol).value, 0))
        If px <= 0 And prevCol > 0 Then px = CDbl(IfZero(wsDash.Cells(found, prevCol).value, 0))
        If px <= 0 Then px = price
        Dim risk As Double
        risk = EstimateOrderRisk(qty, atr, slK, px, slipBp)

        Dim keySess As String: keySess = sessionKey
        Dim usedSess As Double: usedSess = 0
        If sessionRiskUsed.Exists(keySess) Then usedSess = sessionRiskUsed(keySess)
        Dim usedTick As Double: usedTick = 0
        If tickerRiskUsed.Exists(ticker) Then usedTick = tickerRiskUsed(ticker)

        If isClose Then
            usedSess = Application.Max(0, usedSess - risk)
            usedTick = Application.Max(0, usedTick - risk)
            totalRiskUsed = Application.Max(0, totalRiskUsed - risk)
        Else
            usedSess = usedSess + risk
            usedTick = usedTick + risk
            totalRiskUsed = totalRiskUsed + risk
        End If
        sessionRiskUsed(keySess) = usedSess
        tickerRiskUsed(ticker) = usedTick
NextOrder:
    Next r
Done:
    On Error Resume Next
End Sub

Private Sub UpdateRiskUsageFromExecutions()
    On Error GoTo fallback
    If sessionRiskUsed Is Nothing Then Set sessionRiskUsed = CreateObject("Scripting.Dictionary") Else sessionRiskUsed.RemoveAll
    If tickerRiskUsed Is Nothing Then Set tickerRiskUsed = CreateObject("Scripting.Dictionary") Else tickerRiskUsed.RemoveAll
    totalRiskUsed = 0

    Dim wsDash As Worksheet: Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    Dim wsExec As Worksheet: Set wsExec = GetSheet("ExecMon")
    If wsExec Is Nothing Then GoTo fallback

    Dim lastRow As Long
    lastRow = wsExec.Cells(wsExec.Rows.Count, 1).End(xlUp).ROW
    If lastRow < 2 Then Exit Sub

    Dim hdrMap As Object: Set hdrMap = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = wsExec.Cells(1, wsExec.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim key As String
        key = LCase$(Trim$(CStr(wsExec.Cells(1, c).value)))
        If Len(key) > 0 Then hdrMap(key) = c
    Next c

    Dim colDate As Long
    Dim colTicker As Long
    Dim colSide As Long
    Dim colQty As Long
    Dim colPrice As Long

    colDate = FindColumnWithAliases(wsExec, 1, "驍上・・ｮ螢ｽ蠕玖ｭ弱・, "驍上・・ｮ螢ｽ蠕玖脂繝ｻ, "驍上・・ｮ螢ｽ蠕・, "timestamp", "exec_time")
    colTicker = FindColumnWithAliases(wsExec, 1, "鬩ｫ菫ｶ豌帷ｹｧ・ｳ郢晢ｽｼ郢昴・, "鬩ｫ菫ｶ豌・, "symbol", "ticker")
    colSide = FindColumnWithAliases(wsExec, 1, "陞｢・ｲ髮具ｽｷ陋ｹ・ｺ陋ｻ繝ｻ, "陞｢・ｲ髮具ｽｷ", "陋ｹ・ｺ陋ｻ繝ｻ, "side")
    colQty = FindColumnWithAliases(wsExec, 1, "驍上・・ｮ螢ｽ辟夐ｩ･繝ｻ, "隰ｨ・ｰ鬩･繝ｻ, "隴ｬ・ｪ隰ｨ・ｰ", "qty", "quantity")
    colPrice = FindColumnWithAliases(wsExec, 1, "驍上・・ｮ螢ｼ閻ｰ關難ｽ｡", "驍上・・ｮ螢ｻ・ｾ・｡隴ｬ・ｼ", "陷雁・ｽｾ・｡", "price")

    If colDate = 0 Or colTicker = 0 Or colSide = 0 Or colQty = 0 Or colPrice = 0 Then GoTo fallback

    Dim sessionCol As Long: sessionCol = FindColumn(wsDash, DASH_HEADER_ROW, "Session")
    Dim modeCol As Long: modeCol = FindColumn(wsDash, DASH_HEADER_ROW, "SignalMode")
    Dim tCol As Long: tCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    Dim atrCol As Long: atrCol = FindColumn(wsDash, DASH_HEADER_ROW, "ATR_n")
    Dim slkCol As Long: slkCol = FindColumn(wsDash, DASH_HEADER_ROW, "SLk")
    If tCol = 0 Or sessionCol = 0 Or modeCol = 0 Or atrCol = 0 Or slkCol = 0 Then GoTo fallback

    Dim slipBp As Double
    slipBp = CDbl(IfZero(wsDash.Range(DASH_SLIP_BP_CELL).value, DEFAULT_SLIP_BP))

    Dim netQty As Object: Set netQty = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lastRow
        Dim execDate As Variant: execDate = wsExec.Cells(r, colDate).value
        If Not IsDate(execDate) Then GoTo NextExec
        If DateValue(execDate) <> Date Then GoTo NextExec

        Dim ticker As String: ticker = Trim$(CStr(wsExec.Cells(r, colTicker).value))
        If Len(ticker) = 0 Then GoTo NextExec

        Dim qtyVal As Double: qtyVal = CDbl(IfZero(wsExec.Cells(r, colQty).value, 0))
        If qtyVal <= 0 Then GoTo NextExec
        Dim qty As Long: qty = CLng(qtyVal)

        Dim price As Double: price = CDbl(IfZero(wsExec.Cells(r, colPrice).value, 0))
        If price <= 0 Then GoTo NextExec

        Dim sideText As String: sideText = CStr(wsExec.Cells(r, colSide).value)
        Dim sideSign As Long: sideSign = ExecSideSign(sideText)
        If sideSign = 0 Then GoTo NextExec

        Dim dLast As Long: dLast = wsDash.Cells(wsDash.Rows.Count, tCol).End(xlUp).ROW
        Dim dr As Long, found As Long: found = 0
        For dr = DASH_DATA_START To dLast
            If CStr(wsDash.Cells(dr, tCol).value) = ticker Then
                found = dr
                Exit For
            End If
        Next dr
        If found = 0 Then GoTo NextExec

        Dim atr As Double: atr = CDbl(IfZero(wsDash.Cells(found, atrCol).value, 0))
        Dim slK As Double: slK = CDbl(IfZero(wsDash.Cells(found, slkCol).value, 0))
        Dim sessionKey As String: sessionKey = GetSessionKey(CStr(wsDash.Cells(found, sessionCol).value), CStr(wsDash.Cells(found, modeCol).value))

        Dim prevNet As Double
        If netQty.Exists(ticker) Then prevNet = netQty(ticker) Else prevNet = 0
        Dim entryQty As Long, exitQty As Long
        If prevNet = 0 Then
            entryQty = qty
        ElseIf Sgn(prevNet) = sideSign Then
            entryQty = qty
        ElseIf qty <= Abs(prevNet) Then
            exitQty = qty
        Else
            exitQty = CLng(Abs(prevNet))
            entryQty = qty - exitQty
        End If

        Dim newNet As Double
        newNet = prevNet + sideSign * qty
        netQty(ticker) = newNet

        Dim usedSess As Double: If sessionRiskUsed.Exists(sessionKey) Then usedSess = sessionRiskUsed(sessionKey) Else usedSess = 0
        Dim usedTick As Double: If tickerRiskUsed.Exists(ticker) Then usedTick = tickerRiskUsed(ticker) Else usedTick = 0

        If exitQty > 0 Then
            Dim riskExit As Double
            riskExit = EstimateOrderRisk(exitQty, atr, slK, price, slipBp)
            usedSess = Application.Max(0, usedSess - riskExit)
            usedTick = Application.Max(0, usedTick - riskExit)
            totalRiskUsed = Application.Max(0, totalRiskUsed - riskExit)
        End If

        If entryQty > 0 Then
            Dim riskEntry As Double
            riskEntry = EstimateOrderRisk(entryQty, atr, slK, price, slipBp)
            usedSess = usedSess + riskEntry
            usedTick = usedTick + riskEntry
            totalRiskUsed = totalRiskUsed + riskEntry
        End If

        sessionRiskUsed(sessionKey) = usedSess
        tickerRiskUsed(ticker) = usedTick
NextExec:
    Next r
    Exit Sub

fallback:
    RebuildRiskUsageFromOrders
End Sub

Private Sub EnsureOrdersSheet()
    Dim wsOrders As Worksheet
    Set wsOrders = EnsureSheet(SHEET_ORDERS)
    If wsOrders.Cells(1, 1).value = "" Then
        wsOrders.Range("A1:F1").value = Array("Time", "Ticker", "Side", "Price", "Qty", "Note")
    End If
End Sub

Private Sub WriteHeaderTexts(ByVal ws As Worksheet)
    Dim rtHeaders As Variant: rtHeaders = RealTimeHeaderList()
    Dim candHeaders As Variant: candHeaders = CandidateHeaderList()
    Dim baseCol As Long: baseCol = 8 ' column H
    Dim i As Long
    For i = LBound(rtHeaders) To UBound(rtHeaders)
        ws.Cells(DASH_HEADER_ROW, baseCol + i).value = rtHeaders(i)
    Next i

    Dim baseCand As Long
    baseCand = baseCol + UBound(rtHeaders) + 1
    For i = LBound(candHeaders) To UBound(candHeaders)
        ws.Cells(DASH_HEADER_ROW, baseCand + i).value = candHeaders(i)
    Next i
End Sub

Public Sub InstallRealtimeFormulas()
    ' One-time installer to apply formulas safely without clearing existing values.
    EnsureRuntimeReady True
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    ApplyRealtimeColumns ws
    ProtectRealtimeColumns ws
End Sub

Private Sub EnsureHeaders(ByVal ws As Worksheet)
    ' Write labels and ensure formulas are applied once.
    Dim wasProtected As Boolean
    On Error Resume Next
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect Password:=""
    On Error GoTo 0

    WriteHeaderTexts ws
    ApplyRealtimeColumns ws
    ProtectRealtimeColumns ws
End Sub


Private Function DashboardHeadersReady(ByVal ws As Worksheet) As Boolean
    Dim headers As Variant: headers = DashboardHeaderList()
    Dim baseCol As Long: baseCol = 8
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If ws.Cells(DASH_HEADER_ROW, baseCol + i).value <> headers(i) Then
            DashboardHeadersReady = False
            Exit Function
        End If
    Next i
    DashboardHeadersReady = ws.Cells(DASH_DATA_START, baseCol + 1).HasFormula
End Function

Private Sub LoadCandidates()
    Dim path As String
    path = ThisWorkbook.path & CANDIDATES_REL_PATH
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_CANDIDATES)
    ws.Cells.Clear
    If Len(Dir$(path)) = 0 Then
        If IsHeadless() Then
            LogDebug "LoadCandidates missing path=" & path
        Else
            MsgBox "Candidates not found: " & path, vbExclamation
        End If
        Exit Sub
    End If
    Dim f As Integer: f = FreeFile
    Dim line As String
    Dim r As Long: r = 1
    Dim values As Variant
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        values = ParseCsvLine(line)
        ws.Cells(r, 1).Resize(1, UBound(values) + 1).value = values
        r = r + 1
    Loop
    Close #f
    ws.Columns.AutoFit
    NormalizeCandidateHeaders ws
    Dim headless As Boolean
    headless = IsHeadless()
    If headless Then
        LogDebug "LoadCandidates loaded count=" & (r - 2)
    Else
        MsgBox "Loaded " & (r - 2) & " candidates.", vbInformation
    End If
End Sub

Private Sub PushCandidatesToDashboard()
    On Error GoTo FailPush

    Dim wsCand As Worksheet
    Set wsCand = EnsureSheet(SHEET_CANDIDATES)
    Dim lastRow As Long
    lastRow = wsCand.Cells(wsCand.Rows.Count, 1).End(xlUp).ROW
    If lastRow < 2 Then
        If IsHeadless() Then
            LogDebug "PushCandidates empty candidates sheet"
        Else
            MsgBox "Candidates sheet is empty. Load first.", vbExclamation
        End If
        Exit Sub
    End If

    Dim colTicker As Long: colTicker = FindColumnWithAliases(wsCand, 1, "Ticker", "code", "ticker")
    Dim colSel As Long: colSel = FindColumnWithAliases(wsCand, 1, "Selected", "selected", "sel")
    Dim colSignal As Long: colSignal = FindColumnWithAliases(wsCand, 1, "SignalMode", "signal_mode", "mode")
    Dim colSession As Long: colSession = FindColumnWithAliases(wsCand, 1, "Session", "session", "session_label")
    Dim colPlanTag As Long: colPlanTag = FindColumnWithAliases(wsCand, 1, "PlanTag", "plan_tag", "plan")
    Dim colATR As Long: colATR = FindColumn(wsCand, 1, "ATR_n")
    Dim colTP As Long: colTP = FindColumn(wsCand, 1, "TPk")
    Dim colSL As Long: colSL = FindColumn(wsCand, 1, "SLk")
    Dim colJth As Long: colJth = FindColumn(wsCand, 1, "J_th")
    Dim colFpf As Long: colFpf = FindColumnWithAliases(wsCand, 1, "forward_pf_eff", "ForwardPF")
    Dim colFtr As Long: colFtr = FindColumnWithAliases(wsCand, 1, "forward_trades", "ForwardTrades")
    Dim colWinRate As Long: colWinRate = FindColumnWithAliases(wsCand, 1, "forward_winrate", "ForwardWin")
    Dim colWinLow As Long: colWinLow = FindColumnWithAliases(wsCand, 1, "forward_win_ci_low", "WinCI_L")
    Dim colWinHigh As Long: colWinHigh = FindColumnWithAliases(wsCand, 1, "forward_win_ci_high", "WinCI_H")
    Dim colExpMean As Long: colExpMean = FindColumnWithAliases(wsCand, 1, "forward_exp_boot_mean", "ExpBootMean")
    Dim colExpLow As Long: colExpLow = FindColumnWithAliases(wsCand, 1, "forward_exp_boot_low", "ExpBootLow")
    Dim colExpHigh As Long: colExpHigh = FindColumnWithAliases(wsCand, 1, "forward_exp_boot_high", "ExpBootHigh")
    Dim colMaxDD As Long: colMaxDD = FindColumn(wsCand, 1, "MaxDD")
    If colMaxDD = 0 Then colMaxDD = FindColumn(wsCand, 1, "max_dd")
    Dim colForwardAvgBars As Long: colForwardAvgBars = FindColumn(wsCand, 1, "ForwardAvgBars")
    If colForwardAvgBars = 0 Then colForwardAvgBars = FindColumn(wsCand, 1, "forward_avg_bars")
    Dim colGapBucket As Long: colGapBucket = FindColumn(wsCand, 1, "GapBucket")
    If colGapBucket = 0 Then colGapBucket = FindColumn(wsCand, 1, "forward_gap_best_bucket")
    Dim colGapRule As Long: colGapRule = FindColumn(wsCand, 1, "GapRule")
    If colGapRule = 0 Then colGapRule = FindColumn(wsCand, 1, "forward_gap_rule")
    Dim colGapSummary As Long: colGapSummary = FindColumnWithAliases(wsCand, 1, "GapSummary", "gap_summary")
    If colGapSummary = 0 Then colGapSummary = FindColumn(wsCand, 1, "forward_gap_summary")
    If colTicker = 0 Then
        If IsHeadless() Then
            LogDebug "PushCandidates ticker column missing"
        Else
            MsgBox "Ticker column missing in candidates.", vbCritical
        End If
        Exit Sub
    End If

    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    Dim candHeaders As Variant: candHeaders = CandidateHeaderList()
    Dim baseCol As Long: baseCol = FindColumn(wsDash, DASH_HEADER_ROW, "TickerSrc")
    If baseCol = 0 Then baseCol = 8 + UBound(RealTimeHeaderList()) + 1
    Dim clearWidth As Long: clearWidth = UBound(candHeaders)
    wsDash.Range(wsDash.Cells(DASH_DATA_START, baseCol), wsDash.Cells(wsDash.Rows.Count, baseCol + clearWidth)).ClearContents

    Dim colTickerSrcDash As Long: colTickerSrcDash = FindColumn(wsDash, DASH_HEADER_ROW, "TickerSrc")
    Dim colSelectedDash As Long: colSelectedDash = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim colSignalDash As Long: colSignalDash = FindColumn(wsDash, DASH_HEADER_ROW, "SignalMode")
    Dim colSessionDash As Long: colSessionDash = FindColumn(wsDash, DASH_HEADER_ROW, "Session")
    Dim colATRDash As Long: colATRDash = FindColumn(wsDash, DASH_HEADER_ROW, "ATR_n")
    Dim colTPDash As Long: colTPDash = FindColumn(wsDash, DASH_HEADER_ROW, "TPk")
    Dim colSLDash As Long: colSLDash = FindColumn(wsDash, DASH_HEADER_ROW, "SLk")
    Dim colJthDash As Long: colJthDash = FindColumn(wsDash, DASH_HEADER_ROW, "J_th")
    Dim colForwardPFDash As Long: colForwardPFDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardPF")
    Dim colForwardTradesDash As Long: colForwardTradesDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardTrades")
    Dim colForwardWinDash As Long: colForwardWinDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardWin")
    Dim colWinLowDash As Long: colWinLowDash = FindColumn(wsDash, DASH_HEADER_ROW, "WinCI_L")
    Dim colWinHighDash As Long: colWinHighDash = FindColumn(wsDash, DASH_HEADER_ROW, "WinCI_H")
    Dim colExpMeanDash As Long: colExpMeanDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootMean")
    Dim colMaxDDDash As Long: colMaxDDDash = FindColumn(wsDash, DASH_HEADER_ROW, "MaxDD")
    Dim colExpLowDash As Long: colExpLowDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootLow")
    Dim colExpHighDash As Long: colExpHighDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootHigh")
    Dim colForwardAvgDash As Long: colForwardAvgDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardAvgBars")
    Dim colGapBucketDash As Long: colGapBucketDash = FindColumn(wsDash, DASH_HEADER_ROW, "GapBucket")
    Dim colGapRuleDash As Long: colGapRuleDash = FindColumn(wsDash, DASH_HEADER_ROW, "GapRule")
    Dim colGapSummaryDash As Long: colGapSummaryDash = FindColumn(wsDash, DASH_HEADER_ROW, "GapSummary")
    Dim colDynamicQtyDash As Long: colDynamicQtyDash = FindColumn(wsDash, DASH_HEADER_ROW, "DynamicQty")
    Dim colPlanTagDash As Long: colPlanTagDash = FindColumn(wsDash, DASH_HEADER_ROW, "PlanTag")

    Dim r As Long, targetRow As Long
    targetRow = DASH_DATA_START
    Dim selDef As Long: selDef = CLng(IfZero(wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value, DEFAULT_SELECTED_DEFAULT))

    For r = 2 To lastRow
        Dim ticker As String: ticker = CStr(wsCand.Cells(r, colTicker).value)
        If Len(ticker) = 0 Then GoTo NextCandidate

        If colTickerSrcDash > 0 Then wsDash.Cells(targetRow, colTickerSrcDash).value = ticker
        Dim selVal As Variant
        If colSel > 0 Then
            selVal = wsCand.Cells(r, colSel).value
        Else
            selVal = ""
        End If
        If selVal = "" Then selVal = selDef
        If colSelectedDash > 0 Then wsDash.Cells(targetRow, colSelectedDash).value = selVal

        If colSignal > 0 And colSignalDash > 0 Then wsDash.Cells(targetRow, colSignalDash).value = wsCand.Cells(r, colSignal).value
        Dim sessionRaw As String
        If colSession > 0 Then sessionRaw = CStr(wsCand.Cells(r, colSession).value)
        Dim planTagRaw As String
        If colPlanTag > 0 Then planTagRaw = CStr(wsCand.Cells(r, colPlanTag).value)
        Dim sessionLabel As String
        sessionLabel = ResolveSessionLabel(sessionRaw, planTagRaw)
        Dim jthVal As Double
        If colJth > 0 Then jthVal = CDbl(IfZero(wsCand.Cells(r, colJth).value, 0))
        If colJth > 0 Then
            If Abs(jthVal) > 0 And Abs(jthVal) < DASH_MIN_JTH Then GoTo NextCandidate
        End If
        If colSessionDash > 0 Then wsDash.Cells(targetRow, colSessionDash).value = sessionLabel
        If colATR > 0 And colATRDash > 0 Then wsDash.Cells(targetRow, colATRDash).value = wsCand.Cells(r, colATR).value
        If colTP > 0 And colTPDash > 0 Then wsDash.Cells(targetRow, colTPDash).value = wsCand.Cells(r, colTP).value
        If colSL > 0 And colSLDash > 0 Then wsDash.Cells(targetRow, colSLDash).value = wsCand.Cells(r, colSL).value
        If colJth > 0 And colJthDash > 0 Then wsDash.Cells(targetRow, colJthDash).value = wsCand.Cells(r, colJth).value
        If colFpf > 0 And colForwardPFDash > 0 Then wsDash.Cells(targetRow, colForwardPFDash).value = wsCand.Cells(r, colFpf).value
        If colFtr > 0 And colForwardTradesDash > 0 Then wsDash.Cells(targetRow, colForwardTradesDash).value = wsCand.Cells(r, colFtr).value
        If colWinRate > 0 And colForwardWinDash > 0 Then wsDash.Cells(targetRow, colForwardWinDash).value = wsCand.Cells(r, colWinRate).value
        If colWinLow > 0 And colWinLowDash > 0 Then wsDash.Cells(targetRow, colWinLowDash).value = wsCand.Cells(r, colWinLow).value
        If colWinHigh > 0 And colWinHighDash > 0 Then wsDash.Cells(targetRow, colWinHighDash).value = wsCand.Cells(r, colWinHigh).value
        If colExpMean > 0 And colExpMeanDash > 0 Then wsDash.Cells(targetRow, colExpMeanDash).value = wsCand.Cells(r, colExpMean).value
        If colExpLow > 0 And colExpLowDash > 0 Then wsDash.Cells(targetRow, colExpLowDash).value = wsCand.Cells(r, colExpLow).value
        If colExpHigh > 0 And colExpHighDash > 0 Then wsDash.Cells(targetRow, colExpHighDash).value = wsCand.Cells(r, colExpHigh).value
        If colMaxDD > 0 And colMaxDDDash > 0 Then wsDash.Cells(targetRow, colMaxDDDash).value = wsCand.Cells(r, colMaxDD).value
        If colForwardAvgBars > 0 And colForwardAvgDash > 0 Then wsDash.Cells(targetRow, colForwardAvgDash).value = wsCand.Cells(r, colForwardAvgBars).value
        If colGapBucket > 0 And colGapBucketDash > 0 Then wsDash.Cells(targetRow, colGapBucketDash).value = wsCand.Cells(r, colGapBucket).value
        If colGapRule > 0 And colGapRuleDash > 0 Then wsDash.Cells(targetRow, colGapRuleDash).value = wsCand.Cells(r, colGapRule).value
        If colGapSummary > 0 And colGapSummaryDash > 0 Then wsDash.Cells(targetRow, colGapSummaryDash).value = Left$(CStr(wsCand.Cells(r, colGapSummary).value), 255)
        If colDynamicQtyDash > 0 Then wsDash.Cells(targetRow, colDynamicQtyDash).value = ""
        If colPlanTag > 0 And colPlanTagDash > 0 Then wsDash.Cells(targetRow, colPlanTagDash).value = planTagRaw

        targetRow = targetRow + 1
NextCandidate:
    Next r
    ' As a safety net, ensure forward metrics are attached even if header aliases drifted.
    AttachForwardMetrics wsDash, wsCand, DASH_DATA_START, targetRow - 1
    ResolveTickerSelections wsDash, colTickerSrcDash, colSessionDash, colSignalDash, colSelectedDash, targetRow, selDef
    Application.CalculateFull
    On Error Resume Next
    wsDash.Range(wsDash.Cells(DASH_HEADER_ROW, baseCol), wsDash.Cells(DASH_DATA_START - 1 + (targetRow - DASH_DATA_START), baseCol + clearWidth)).EntireColumn.AutoFit
    If Err.Number <> 0 Then
        LogDebug "PushCandidates autofit skipped err=" & Err.Number & " desc=" & Err.Description
        Err.Clear
    End If
    On Error GoTo FailPush
    If IsHeadless() Then
        LogDebug "PushCandidates completed count=" & (targetRow - DASH_DATA_START)
    Else
        MsgBox "Dashboard updated with " & (targetRow - DASH_DATA_START) & " tickers.", vbInformation
    End If
    Exit Sub

FailPush:
    LogDebug "PushCandidates error err=" & Err.Number & " desc=" & Err.Description
    If Not IsHeadless() Then
        MsgBox "PushCandidates failed: " & Err.Description, vbCritical
    End If
    Err.Clear
End Sub

Private Sub ResolveTickerSelections(ByVal ws As Worksheet, ByVal colTicker As Long, ByVal colSession As Long, ByVal colSignal As Long, ByVal colSelected As Long, ByVal targetRow As Long, ByVal selDefault As Long)
    If colTicker = 0 Or colSession = 0 Or colSignal = 0 Or colSelected = 0 Then Exit Sub
    Dim bestRow As Object: Set bestRow = CreateObject("Scripting.Dictionary")
    Dim bestPriority As Object: Set bestPriority = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = DASH_DATA_START To targetRow - 1
        Dim ticker As String: ticker = CStr(ws.Cells(r, colTicker).value)
        If Len(ticker) = 0 Then GoTo ContinueLoop
        Dim sess As String: sess = CStr(ws.Cells(r, colSession).value)
        Dim mode As String: mode = CStr(ws.Cells(r, colSignal).value)
        Dim pr As Double: pr = GetSessionPriority(sess, mode)
        If Not bestRow.Exists(ticker) Then
            bestRow(ticker) = r
            bestPriority(ticker) = pr
            ws.Cells(r, colSelected).value = selDefault
        ElseIf pr < bestPriority(ticker) - 0.0001 Then
            ws.Cells(bestRow(ticker), colSelected).value = 0
            bestRow(ticker) = r
            bestPriority(ticker) = pr
            ws.Cells(r, colSelected).value = selDefault
        Else
            ws.Cells(r, colSelected).value = 0
        End If
ContinueLoop:
    Next r
End Sub

Private Sub StartAutoTrading()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    wsDash.Range(DASH_STATUS_CELL).value = 1
    If prevJ Is Nothing Then Set prevJ = CreateObject("Scripting.Dictionary")
    prevJ.RemoveAll
    tradeDate = Date
    orderCount = 0
    isRunning = True
    ScheduleNextTick 1
    LogDebug "Auto trading loop started (dry-run)."
End Sub

Private Sub StopAutoTrading()
    On Error GoTo EvalError
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    wsDash.Range(DASH_STATUS_CELL).value = 0
    isRunning = False
    On Error Resume Next
    If AutoTimer <> 0 Then Application.OnTime AutoTimer, "AutoTrader.AutoTick", , False
    CancelCloseExit
    On Error GoTo 0
    LogDebug "Auto trading loop stopped."
    Exit Sub
EvalError:
    LogDebug "StopAutoTrading error " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

Private Sub ScheduleNextTick(ByVal seconds As Double)
    On Error Resume Next
    AutoTimer = Now + TimeSerial(0, 0, seconds)
    Application.OnTime AutoTimer, "AutoTrader.AutoTick"
    On Error GoTo 0
End Sub

Public Sub AutoTick()
    On Error GoTo ExitTick
    If Not isRunning Then Exit Sub
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    If wsDash.Range(DASH_STATUS_CELL).value <> 1 Then
        StopAutoTrading
        Exit Sub
    End If
    If CLng(IfZero(wsDash.Range(DASH_HARDSTOP_CELL).value, 0)) = 1 Then Exit Sub
    If Date <> tradeDate Then
        orderCount = 0
        tradeDate = Date
        If Not prevJ Is Nothing Then prevJ.RemoveAll
        If Not sessionRiskUsed Is Nothing Then sessionRiskUsed.RemoveAll
        If Not tickerRiskUsed Is Nothing Then tickerRiskUsed.RemoveAll
        totalRiskUsed = 0
        UpdateRiskUsageFromExecutions
        If CLng(IfZero(wsDash.Range(DASH_REENTRY_CELL).value, DEFAULT_REENTRY)) = 1 Then
            ResetSelectedToDefault wsDash
        End If
    End If

    EvaluateAndQueueOrders
    ScheduleNextTick 1
ExitTick:
End Sub
Private Sub ResetSelectedToDefault(ByVal wsDash As Worksheet)
    Dim selCol As Long: selCol = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim tickerCol As Long: tickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    If selCol = 0 Or tickerCol = 0 Then Exit Sub
    Dim lastRow As Long: lastRow = wsDash.Cells(wsDash.Rows.Count, tickerCol).End(xlUp).ROW
    Dim defVal As Long: defVal = CLng(IfZero(wsDash.Range(DASH_SELECTED_DEFAULT_CELL).value, DEFAULT_SELECTED_DEFAULT))
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        If wsDash.Cells(r, tickerCol).value <> "" Then
            wsDash.Cells(r, selCol).value = defVal
        End If
    Next r
End Sub

Private Sub EvaluateAndQueueOrders()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHEET_DASHBOARD)
    RepairRealtimeIfCorrupted wsDash
    Dim startTime As Date, endTime As Date
    startTime = ParseTime(wsDash.Range(DASH_SESSION_START_CELL).value, DEFAULT_SESSION_START)
    endTime = ParseTime(wsDash.Range(DASH_SESSION_END_CELL).value, DEFAULT_SESSION_END)
    Dim tm As Date: tm = Time
    If tm < startTime Or tm > endTime Then Exit Sub

    Dim maxOrders As Long
    maxOrders = CLng(IfZero(wsDash.Range(DASH_MAX_ORDERS_CELL).value, DEFAULT_MAX_ORDERS))
    If orderCount >= maxOrders Then Exit Sub

    Dim totalBudget As Double
    totalBudget = CDbl(IfZero(wsDash.Range(DASH_BUDGET_CELL).value, DEFAULT_MAX_BUDGET))
    If totalBudget <= 0 Then totalBudget = DEFAULT_MAX_BUDGET

    If sessionRiskUsed Is Nothing Then Set sessionRiskUsed = CreateObject("Scripting.Dictionary")
    If tickerRiskUsed Is Nothing Then Set tickerRiskUsed = CreateObject("Scripting.Dictionary")
    UpdateRiskUsageFromExecutions

    Dim selCol As Long: selCol = FindColumn(wsDash, DASH_HEADER_ROW, "Selected")
    Dim signalCol As Long: signalCol = FindColumn(wsDash, DASH_HEADER_ROW, "SignalMode")
    Dim sessionCol As Long: sessionCol = FindColumn(wsDash, DASH_HEADER_ROW, "Session")
    Dim tickerCol As Long: tickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    Dim priceCol As Long: priceCol = FindColumn(wsDash, DASH_HEADER_ROW, "PrevClose")
    If priceCol = 0 Then priceCol = FindColumn(wsDash, DASH_HEADER_ROW, "PreOpenMid")
    Dim jCol As Long: jCol = FindColumn(wsDash, DASH_HEADER_ROW, "J")
    Dim jthCol As Long: jthCol = FindColumn(wsDash, DASH_HEADER_ROW, "J_th")
    Dim qtyCol As Long: qtyCol = FindColumn(wsDash, DASH_HEADER_ROW, "DynamicQty")
    Dim signalStatusCol As Long: signalStatusCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderSignalStatusJP())
    Dim signalKindCol As Long: signalKindCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderSignalKindJP())
    Dim colForwardPFDash As Long: colForwardPFDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardPF")
    Dim colForwardTradesDash As Long: colForwardTradesDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardTrades")
    Dim colForwardWinDash As Long: colForwardWinDash = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardWin")
    Dim colExpMeanDash As Long: colExpMeanDash = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootMean")
    Dim colMaxDDDash As Long: colMaxDDDash = FindColumn(wsDash, DASH_HEADER_ROW, "MaxDD")
    Dim colSLDash As Long: colSLDash = FindColumn(wsDash, DASH_HEADER_ROW, "SLk")
    Dim colATRDash As Long: colATRDash = FindColumn(wsDash, DASH_HEADER_ROW, "ATR_n")
    Dim colAtrValueDash As Long: colAtrValueDash = FindColumn(wsDash, DASH_HEADER_ROW, "ATR")
    If priceCol = 0 Or selCol = 0 Or signalCol = 0 Or sessionCol = 0 Or tickerCol = 0 Or jCol = 0 Then Exit Sub
    If colForwardPFDash = 0 Or colForwardTradesDash = 0 Or colForwardWinDash = 0 Or colExpMeanDash = 0 Then Exit Sub
    If colSLDash = 0 Then Exit Sub
    If colATRDash = 0 Then colATRDash = FindColumn(wsDash, DASH_HEADER_ROW, "ATR_n")

    Dim defaultQty As Long
    defaultQty = CLng(IfZero(wsDash.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))
    Dim lotStep As Long
    lotStep = CLng(IfZero(wsDash.Range(DASH_LOT_STEP_CELL).value, DEFAULT_LOT_STEP))
    Dim slipBp As Double
    slipBp = CDbl(IfZero(wsDash.Range(DASH_SLIP_BP_CELL).value, DEFAULT_SLIP_BP))

    Dim rowScores As Object: Set rowScores = CreateObject("Scripting.Dictionary")
    Dim sessionScores As Object: Set sessionScores = CreateObject("Scripting.Dictionary")

    Dim ticker As String
    Dim sessionVal As String
    Dim modeVal As String
    Dim sessionKey As String
    Dim weight As Double
    Dim winRate As Double
    Dim pfEff As Double
    Dim tradesVal As Double
    Dim expMean As Double
    Dim scoreVal As Double
    Dim ddVal As Double
    Dim rowScore As Double
    Dim px As Double
    Dim allocBudget As Double
    Dim sessionCap As Double
    Dim sessionUsed As Double
    Dim sessionAvail As Double
    Dim capFraction As Double
    Dim orderCap As Double
    Dim side As String
    Dim qty As Long
    Dim jVal As Double
    Dim mode As String
    Dim sessionScoreSum As Double
    Dim sessionWeight As Double
    Dim atrParam As Double
    Dim atrValue As Double
    Dim slVal As Double
    Dim riskVal As Double

    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, tickerCol).End(xlUp).ROW
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        rowScores(CStr(r)) = 0#
        Dim selVal As Variant: selVal = wsDash.Cells(r, selCol).value
        If selVal <> 1 Then GoTo NextScore
        ticker = CStr(wsDash.Cells(r, tickerCol).value)
        If Len(ticker) = 0 Then GoTo NextScore
        sessionVal = CStr(wsDash.Cells(r, sessionCol).value)
        modeVal = CStr(wsDash.Cells(r, signalCol).value)
        sessionKey = GetSessionKey(sessionVal, modeVal)
        weight = GetSessionWeight(sessionKey)
        If weight <= 0 Then
            wsDash.Cells(r, selCol).value = 0
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "SKIP_SESSION"
            GoTo NextScore
        End If
        winRate = CDbl(IfZero(wsDash.Cells(r, colForwardWinDash).value, 0))
        pfEff = CDbl(IfZero(wsDash.Cells(r, colForwardPFDash).value, 0))
        tradesVal = CDbl(IfZero(wsDash.Cells(r, colForwardTradesDash).value, 0))
        expMean = CDbl(IfZero(wsDash.Cells(r, colExpMeanDash).value, 0))
        If colMaxDDDash > 0 Then
            ddVal = CDbl(IfZero(wsDash.Cells(r, colMaxDDDash).value, 0))
        Else
            ddVal = 0
        End If
        scoreVal = ComputeRowScore(winRate, pfEff, tradesVal, expMean, ddVal)
        If scoreVal <= 0 Then
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "SKIP_SCORE"
            GoTo NextScore
        End If
        rowScores(CStr(r)) = scoreVal
        If sessionScores.Exists(sessionKey) Then
            sessionScores(sessionKey) = sessionScores(sessionKey) + scoreVal
        Else
            sessionScores(sessionKey) = scoreVal
        End If
NextScore:
    Next r

    For r = DASH_DATA_START To lastRow
        ticker = CStr(wsDash.Cells(r, tickerCol).value)
        jVal = CDbl(IfZero(wsDash.Cells(r, jCol).value, 0))
        mode = LCase$(CStr(wsDash.Cells(r, signalCol).value))
        sessionVal = CStr(wsDash.Cells(r, sessionCol).value)
        If Len(ticker) = 0 Then GoTo UpdatePrev
        If wsDash.Cells(r, selCol).value <> 1 Then
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = ""
            If signalKindCol > 0 Then wsDash.Cells(r, signalKindCol).value = ""
            If qtyCol > 0 Then wsDash.Cells(r, qtyCol).value = ""
            GoTo UpdatePrev
        End If

        If rowScores.Exists(CStr(r)) Then rowScore = rowScores(CStr(r))
        If rowScore <= 0 Then
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "SKIP"
            wsDash.Cells(r, selCol).value = 0
            If qtyCol > 0 Then wsDash.Cells(r, qtyCol).value = ""
            GoTo UpdatePrev
        End If

        sessionKey = GetSessionKey(sessionVal, mode)
        sessionWeight = GetSessionWeight(sessionKey)
        If sessionWeight <= 0 Then
            wsDash.Cells(r, selCol).value = 0
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "SKIP_SESSION"
            GoTo UpdatePrev
        End If

        If sessionScores.Exists(sessionKey) Then sessionScoreSum = sessionScores(sessionKey)
        If sessionScoreSum <= 0 Then GoTo UpdatePrev

        ' base price will be selected after side is determined

        allocBudget = totalBudget * sessionWeight * (rowScore / sessionScoreSum)
        sessionCap = totalBudget * sessionWeight
        If sessionRiskUsed.Exists(sessionKey) Then sessionUsed = sessionRiskUsed(sessionKey)
        sessionAvail = sessionCap - sessionUsed
        If sessionAvail <= 0 Then
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "SESSION_FULL"
            wsDash.Cells(r, selCol).value = 0
            GoTo UpdatePrev
        End If
        If allocBudget > sessionAvail Then allocBudget = sessionAvail
        capFraction = GetOrderCapFraction(sessionKey)
        If capFraction > 0 Then
            orderCap = totalBudget * capFraction
            If allocBudget > orderCap Then allocBudget = orderCap
        End If
        If allocBudget <= 0 Then GoTo UpdatePrev

        If jVal < 0 Then
            side = "BUY"
        Else
            side = "SELL"
        End If
        ' choose side-specific quote first, then fall back to mid/prev
        Dim bidCol As Long: bidCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderPreopenBidJP())
        Dim askCol As Long: askCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderPreopenAskJP())
        Dim midCol As Long: midCol = FindColumn(wsDash, DASH_HEADER_ROW, HeaderPreopenMidJP())
        Dim pxBid As Double, pxAsk As Double, pxMid As Double, pxPrev As Double
        If bidCol > 0 Then pxBid = CDbl(IfZero(wsDash.Cells(r, bidCol).value, 0))
        If askCol > 0 Then pxAsk = CDbl(IfZero(wsDash.Cells(r, askCol).value, 0))
        If midCol > 0 Then pxMid = CDbl(IfZero(wsDash.Cells(r, midCol).value, 0))
        If priceCol > 0 Then pxPrev = CDbl(IfZero(wsDash.Cells(r, priceCol).value, 0))
        If side = "BUY" Then
            px = IIf(pxAsk > 0, pxAsk, IIf(pxMid > 0, pxMid, pxPrev))
        Else
            px = IIf(pxBid > 0, pxBid, IIf(pxMid > 0, pxMid, pxPrev))
        End If
        If px <= 0 Then
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "NO_PRICE"
            GoTo UpdatePrev
        End If

        ' Prefer sheet-driven DynamicQty when present; fallback to VBA sizing
        Dim qtyFromSheet As Long
        qtyFromSheet = 0
        If qtyCol > 0 Then
            On Error Resume Next
            If wsDash.Cells(r, qtyCol).HasFormula Then
                qtyFromSheet = CLng(IfZero(wsDash.Cells(r, qtyCol).value, 0))
            Else
                qtyFromSheet = CLng(IfZero(wsDash.Cells(r, qtyCol).value, 0))
            End If
            On Error GoTo 0
        End If
        If qtyFromSheet > 0 Then
            qty = qtyFromSheet
        Else
            qty = ComputeDynamicQty(px, side, allocBudget, lotStep, slipBp, defaultQty)
        End If
        If qty <= 0 Then
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "NO_QTY"
            wsDash.Cells(r, selCol).value = 0
            GoTo UpdatePrev
        End If

        atrParam = CDbl(IfZero(wsDash.Cells(r, colATRDash).value, 0))
        If colAtrValueDash > 0 Then
            atrValue = CDbl(IfZero(wsDash.Cells(r, colAtrValueDash).value, atrParam))
        Else
            atrValue = atrParam
        End If
        slVal = CDbl(IfZero(wsDash.Cells(r, colSLDash).value, 0))
        riskVal = EstimateOrderRisk(qty, atrValue, slVal, px, slipBp)
        If Not CheckRiskAndBudgetLimits(ticker, sessionKey, riskVal, totalBudget) Then
            wsDash.Cells(r, selCol).value = 0
            If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "RISK_LIMIT"
            GoTo UpdatePrev
        End If

        If qtyCol > 0 Then
            If Not wsDash.Cells(r, qtyCol).HasFormula Then
                wsDash.Cells(r, qtyCol).value = qty
            End If
        End If
        If signalStatusCol > 0 Then wsDash.Cells(r, signalStatusCol).value = "ORDERED"
        If signalKindCol > 0 Then wsDash.Cells(r, signalKindCol).value = side & " / " & mode
        PlaceOrder ticker, side, px, qty, mode & ":" & sessionVal
        PlaceBracketIfAvailable wsDash, r, ticker, side, px, qty
        ScheduleCloseExit wsDash, ticker, side
        wsDash.Cells(r, selCol).value = 0
        orderCount = orderCount + 1
        If orderCount >= maxOrders Then Exit For
UpdatePrev:
        SetPrevJ ticker, CDbl(IfZero(wsDash.Cells(r, jCol).value, 0))
    Next r
    Exit Sub
EvalError:
    LogDebug "EvaluateAndQueueOrders error " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

Private Function ResolveSessionLabel(ByVal sessionRaw As String, ByVal planTag As String) As String
    Dim token As String
    Dim tagBase As String
    tagBase = Trim$(planTag)
    If InStr(1, tagBase, "_") > 0 Then tagBase = Split(tagBase, "_")(0)
    token = NormalizeSessionToken(tagBase)
    If Len(token) = 0 Then
        token = NormalizeSessionToken(sessionRaw)
    End If
    If Len(token) = 0 Then
        ResolveSessionLabel = ""
    Else
        ResolveSessionLabel = FormatSessionDisplay(token)
    End If
End Function

Private Function NormalizeSessionToken(ByVal raw As String) As String
    Dim text As String
    text = CStr(raw)
    Dim cleaned As String
    cleaned = Replace(text, ChrW$(&H3000), " ")
    cleaned = Replace(cleaned, ":", "")
    cleaned = Replace(cleaned, "-", "")
    If InStr(cleaned, "_") > 0 Then cleaned = Split(cleaned, "_")(0)
    cleaned = UCase$(Trim$(cleaned))
    Select Case cleaned
        Case "", "N/A", "NA"
            NormalizeSessionToken = ""
            Exit Function
    End Select
    Select Case cleaned
        Case "AM15", "AM0915", "AM915", "AM9", "AM09", "AM09 15"
            NormalizeSessionToken = "AM0915"
            Exit Function
        Case "AM0930", "AM930", "AM9 30"
            NormalizeSessionToken = "AM0930"
            Exit Function
        Case "AM0945", "AM945", "AM9 45"
            NormalizeSessionToken = "AM0945"
            Exit Function
        Case "AM10", "AM1000", "AM100", "AM10 00"
            NormalizeSessionToken = "AM1000"
            Exit Function
        Case "AM1015", "AM10 15"
            NormalizeSessionToken = "AM1015"
            Exit Function
        Case "AM1030", "AM10 30"
            NormalizeSessionToken = "AM1030"
            Exit Function
        Case "PM1430", "PM14 30", "PM143"
            NormalizeSessionToken = "PM1430"
            Exit Function
    End Select
    If Len(cleaned) > 2 Then
        Dim prefix As String: prefix = Left$(cleaned, 2)
        Dim digits As String: digits = Mid$(cleaned, 3)
        digits = Replace(digits, " ", "")
        If Len(digits) = 3 Then digits = "0" & digits
        If Len(digits) = 2 Then digits = "0" & digits & "0"
        If Len(digits) = 4 Then
            NormalizeSessionToken = prefix & digits
            Exit Function
        End If
    End If
    NormalizeSessionToken = ""
End Function

Private Function FormatSessionDisplay(ByVal token As String) As String
    If Len(token) <> 6 Then
        FormatSessionDisplay = token
    Else
        FormatSessionDisplay = Left$(token, 2) & Mid$(token, 3, 2) & ":" & Right$(token, 2)
    End If
End Function

Private Function GetSessionKey(ByVal sessionVal As String, ByVal modeVal As String) As String
    Dim token As String
    token = NormalizeSessionToken(sessionVal)
    If Len(token) = 0 Then token = "UNKNOWN"
    Dim modeKey As String
    modeKey = LCase$(Trim$(modeVal))
    If Len(modeKey) = 0 Then modeKey = "unknown"
    GetSessionKey = token & "::" & modeKey
End Function

Private Function GetSessionWeight(ByVal sessionKey As String) As Double
    Dim parts() As String
    parts = Split(sessionKey, "::")
    Dim token As String: token = parts(0)
    Dim baseWeight As Double
    Select Case token
        Case "AM0915"
            baseWeight = 0
        Case "AM0930"
            baseWeight = 0.2
        Case "AM0945"
            baseWeight = 0.18
        Case "AM1000"
            baseWeight = 0.15
        Case "AM1015"
            baseWeight = 0.14
        Case "AM1030"
            baseWeight = 0.12
        Case "PM1430", "AM15"
            baseWeight = 0
        Case Else
            baseWeight = 0.08
    End Select
    Dim modeRatio As Double
    Dim modeKey As String
    If UBound(parts) >= 1 Then modeKey = parts(1)
    Select Case modeKey
        Case "j-only", "j_only", "jonly"
            modeRatio = SESSION_MODE_JONLY_RATIO
        Case "j-cross", "j_cross", "jcross"
            modeRatio = SESSION_MODE_JCROSS_RATIO
        Case Else
            modeRatio = 0.5
    End Select
    GetSessionWeight = baseWeight * modeRatio
End Function

Private Function GetSessionPriority(ByVal sessionVal As String, ByVal modeVal As String) As Double
    Dim token As String
    token = NormalizeSessionToken(sessionVal)
    Dim base As Double
    Select Case token
        Case "AM0915"
            base = 1
        Case "AM0930"
            base = 2
        Case "AM0945"
            base = 3
        Case "AM1000"
            base = 4
        Case "AM1015"
            base = 5
        Case "AM1030"
            base = 6
        Case "PM1430"
            base = 7
        Case Else
            base = 50
    End Select
    Dim modeAdj As Double
    Select Case LCase$(Trim$(modeVal))
        Case "j-only", "j_only", "jonly"
            modeAdj = -0.2
        Case "j-cross", "j_cross", "jcross"
            modeAdj = -0.1
        Case Else
            modeAdj = 0
    End Select
    GetSessionPriority = base + modeAdj
End Function

Private Function GetOrderCapFraction(ByVal sessionKey As String) As Double
    Dim token As String
    token = Split(sessionKey, "::")(0)
    Select Case token
        Case "AM0915"
            GetOrderCapFraction = 0.07
        Case "AM0930"
            GetOrderCapFraction = 0.06
        Case "AM0945"
            GetOrderCapFraction = 0.055
        Case "AM1000"
            GetOrderCapFraction = 0.05
        Case "AM1015"
            GetOrderCapFraction = 0.045
        Case "AM1030"
            GetOrderCapFraction = 0.04
        Case Else
            GetOrderCapFraction = 0.05
    End Select
End Function

Private Function ComputeRowScore(ByVal winRate As Double, ByVal pfEff As Double, ByVal tradesVal As Double, ByVal expMean As Double, Optional ByVal maxDD As Double = 0) As Double
    If winRate <= 0 Or pfEff <= 0 Or tradesVal <= 0 Then
        ComputeRowScore = 0
        Exit Function
    End If
    Dim tradesFactor As Double
    tradesFactor = tradesVal / ROW_SCORE_TRADE_TARGET
    If tradesFactor > 1 Then tradesFactor = 1
    Dim expFactor As Double
    If expMean <= 0 Then
        expFactor = 0
    Else
        expFactor = expMean / ROW_SCORE_EXP_TARGET
        If expFactor > 1 Then expFactor = 1
    End If
    Dim pfCap As Double
    pfCap = 1# + (tradesVal / ROW_SCORE_TRADE_TARGET) * ROW_SCORE_PF_CAP_GROWTH
    If pfCap < 1.5 Then pfCap = 1.5
    Dim pfFactor As Double
    pfFactor = pfEff
    If pfFactor > pfCap Then pfFactor = pfCap
    If pfFactor < 1 Then pfFactor = 0
    Dim ddNorm As Double
    ddNorm = maxDD
    If ddNorm < 0 Then ddNorm = 0
    If ddNorm > 1 Then ddNorm = ddNorm / 100
    If ddNorm > 1 Then ddNorm = 1
    Dim ddPenalty As Double
    ddPenalty = 1# / (1# + ROW_SCORE_DD_K * ddNorm)
    If ddPenalty < 0.25 Then ddPenalty = 0.25
    ComputeRowScore = winRate * pfFactor * (0.5 + tradesFactor / 2) * (0.5 + expFactor / 2) * ddPenalty
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        On Error Resume Next
        ws.name = name
        On Error GoTo 0
    End If
    Set EnsureSheet = ws
End Function

Private Function GetSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function QuoteForFormula(ByVal s As String) As String
    QuoteForFormula = """" & Replace(CStr(s), """", """""") & """"
End Function

Private Function EstimateOrderRisk(ByVal qty As Long, ByVal atrValue As Double, ByVal slVal As Double, ByVal price As Double, ByVal slipBp As Double) As Double
    If qty <= 0 Then
        EstimateOrderRisk = 0
        Exit Function
    End If

    Dim stopDistance As Double
    stopDistance = atrValue * slVal
    If stopDistance < 0 Then stopDistance = 0

    Dim slipAmount As Double
    slipAmount = Abs(slipBp) / 10000# * price

    Dim perShareRisk As Double
    perShareRisk = stopDistance + slipAmount
    If perShareRisk <= 0 Then perShareRisk = Abs(price) * 0.01

    EstimateOrderRisk = qty * perShareRisk
End Function

Private Function ExecSideSign(ByVal text As String) As Long
    Dim s As String
    s = Trim$(UCase$(CStr(text)))
    If Len(s) = 0 Then
        ExecSideSign = 0
        Exit Function
    End If
    If InStr(s, "鬯ｮ・ｮ陷茨ｽｷ繝ｻ・ｽ繝ｻ・ｷ") > 0 Or Left$(s, 1) = "B" Then
        ExecSideSign = 1
    ElseIf InStr(s, "鬮ｯ讖ｸ・ｽ・｢郢晢ｽｻ繝ｻ・ｲ") > 0 Or Left$(s, 1) = "S" Or Left$(s, 1) = "F" Then
        ExecSideSign = -1
    Else
        ExecSideSign = 0
    End If
End Function

Private Function IfZero(ByVal value As Variant, ByVal fallback As Variant) As Variant
    If IsError(value) Then
        IfZero = fallback
    ElseIf IsEmpty(value) Then
        IfZero = fallback
    ElseIf value = "" Then
        IfZero = fallback
    Else
        IfZero = value
    End If
End Function

Private Function FindColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal name As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim target As String
    target = LCase$(Trim$(name))
    Dim c As Long
    For c = 1 To lastCol
        Dim text As String
        text = LCase$(Trim$(CStr(ws.Cells(headerRow, c).value)))
        text = Replace(text, ChrW$(&HFEFF), "")
        If text = target Then
            FindColumn = c
            Exit Function
        End If
    Next c
    FindColumn = 0
End Function
Private Function ParseCsvLine(ByVal line As String) As Variant
    Dim result() As String
    Dim buffer As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim ch As String
    Dim length As Long

    buffer = ""
    inQuotes = False
    length = Len(line)

    If length = 0 Then
        ReDim result(0)
        result(0) = ""
        ParseCsvLine = result
        Exit Function
    End If

    For i = 1 To length
        ch = Mid$(line, i, 1)
        If ch = """" Then
            If inQuotes And i < length And Mid$(line, i + 1, 1) = """" Then
                buffer = buffer & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            Call AppendCsvValue(result, buffer)
            buffer = ""
        Else
            buffer = buffer & ch
        End If
    Next i

    Call AppendCsvValue(result, buffer)

    ParseCsvLine = result
End Function

Private Sub AppendCsvValue(ByRef arr() As String, ByVal value As String)
    Dim cleaned As String
    cleaned = CleanCsvValue(value)
    Dim idx As Long
    If Not Not arr Then
        idx = UBound(arr) + 1
        ReDim Preserve arr(idx)
        arr(idx) = cleaned
    Else
        ReDim arr(0)
        arr(0) = cleaned
    End If
End Sub
Private Function CleanCsvValue(ByVal value As String) As String
    Dim cleaned As String
    cleaned = value
    cleaned = Replace(cleaned, ChrW$(&HFEFF), "")
    cleaned = Replace(cleaned, ChrW$(&HFFFD), "")
    cleaned = Replace(cleaned, ChrW$(&H202A), "")
    cleaned = Replace(cleaned, ChrW$(&H202C), "")
    CleanCsvValue = Trim$(cleaned)
End Function

Private Sub ProtectRealtimeColumns(ByVal ws As Worksheet)
    On Error Resume Next
    Dim firstCol As Long: firstCol = 8  ' H
    Dim lastCol As Long: lastCol = 20   ' T
    Dim firstRow As Long: firstRow = DASH_DATA_START
    Dim lastRow As Long: lastRow = DASH_DATA_START + DASH_FORMULA_ROWS
    ' Lock only realtime block; other cells remain editable
    ws.Cells.Locked = False
    ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).Locked = True
    ' Allow VBA edits but block user edits on UI
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowSorting:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub

Private Sub RepairRealtimeIfCorrupted(ByVal ws As Worksheet)
    On Error Resume Next
    Dim c As Long
    For c = 9 To 14 ' I..N
        Dim hasF As Boolean
        hasF = ws.Cells(DASH_DATA_START, c).HasFormula
        If Not hasF Then
            ' Reinstall full column formulas for safety
            InstallRealtimeFormulas
            Exit Sub
        End If
    Next c
    On Error GoTo 0
End Sub

' --- Helpers added to complete module and fix compile/runtime errors ---

Private Sub NormalizeCandidateHeaders(ByVal ws As Worksheet)
    ' Normalize header row (row 1) to canonical names used by PushCandidatesToDashboard
    Const ROW As Long = 1
    Dim lastCol As Long
    lastCol = ws.Cells(ROW, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim raw As String
        raw = LCase$(Trim$(CStr(ws.Cells(ROW, c).value)))
        If Len(raw) = 0 Then GoTo NextC
        Select Case raw
            Case "code", "ticker"
                ws.Cells(ROW, c).value = "Ticker"
            Case "selected", "sel"
                ws.Cells(ROW, c).value = "Selected"
            Case "signal_mode", "mode"
                ws.Cells(ROW, c).value = "SignalMode"
            Case "session", "session_label"
                ws.Cells(ROW, c).value = "Session"
            Case "plan_tag", "plan"
                ws.Cells(ROW, c).value = "PlanTag"
            Case "atr_n", "atr"
                ws.Cells(ROW, c).value = "ATR_n"
            Case "tpk", "tp_k"
                ws.Cells(ROW, c).value = "TPk"
            Case "slk", "sl_k"
                ws.Cells(ROW, c).value = "SLk"
            Case "j_th", "jth"
                ws.Cells(ROW, c).value = "J_th"
            Case "forward_pf_eff"
                ws.Cells(ROW, c).value = "ForwardPF"
            Case "forward_trades"
                ws.Cells(ROW, c).value = "ForwardTrades"
            Case "forward_winrate"
                ws.Cells(ROW, c).value = "ForwardWin"
            Case "forward_win_ci_low"
                ws.Cells(ROW, c).value = "WinCI_L"
            Case "forward_win_ci_high"
                ws.Cells(ROW, c).value = "WinCI_H"
            Case "forward_exp_boot_mean"
                ws.Cells(ROW, c).value = "ExpBootMean"
            Case "forward_exp_boot_low"
                ws.Cells(ROW, c).value = "ExpBootLow"
            Case "forward_exp_boot_high"
                ws.Cells(ROW, c).value = "ExpBootHigh"
            Case "maxdd", "max_dd"
                ws.Cells(ROW, c).value = "MaxDD"
            Case "forward_avg_bars"
                ws.Cells(ROW, c).value = "ForwardAvgBars"
            Case "gapbucket", "forward_gap_best_bucket"
                ws.Cells(ROW, c).value = "GapBucket"
            Case "gaprule", "forward_gap_rule"
                ws.Cells(ROW, c).value = "GapRule"
            Case "gapsummary", "gap_summary", "forward_gap_summary"
                ws.Cells(ROW, c).value = "GapSummary"
        End Select
NextC:
    Next c
End Sub

Private Sub AttachForwardMetrics(ByVal wsDash As Worksheet, ByVal wsCand As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next
    If lastRow < firstRow Then Exit Sub
    Dim dTickerCol As Long: dTickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "TickerSrc")
    If dTickerCol = 0 Then dTickerCol = FindColumn(wsDash, DASH_HEADER_ROW, "Ticker")
    If dTickerCol = 0 Then Exit Sub

    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim cTickerCol As Long: cTickerCol = FindColumnWithAliases(wsCand, 1, "Ticker", "code", "ticker")
    If cTickerCol = 0 Then Exit Sub
    Dim cLast As Long: cLast = wsCand.Cells(wsCand.Rows.Count, cTickerCol).End(xlUp).ROW
    Dim cr As Long
    For cr = 2 To cLast
        Dim tk As String: tk = CStr(wsCand.Cells(cr, cTickerCol).value)
        If Len(tk) > 0 Then If Not map.Exists(tk) Then map(tk) = cr
    Next cr

    Dim dashCols As Object: Set dashCols = CreateObject("Scripting.Dictionary")
    dashCols("ForwardPF") = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardPF")
    dashCols("ForwardTrades") = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardTrades")
    dashCols("ForwardWin") = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardWin")
    dashCols("WinCI_L") = FindColumn(wsDash, DASH_HEADER_ROW, "WinCI_L")
    dashCols("WinCI_H") = FindColumn(wsDash, DASH_HEADER_ROW, "WinCI_H")
    dashCols("ExpBootMean") = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootMean")
    dashCols("ExpBootLow") = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootLow")
    dashCols("ExpBootHigh") = FindColumn(wsDash, DASH_HEADER_ROW, "ExpBootHigh")
    dashCols("ForwardAvgBars") = FindColumn(wsDash, DASH_HEADER_ROW, "ForwardAvgBars")

    Dim candCols As Object: Set candCols = CreateObject("Scripting.Dictionary")
    candCols("ForwardPF") = FindColumnWithAliases(wsCand, 1, "ForwardPF", "forward_pf_eff")
    candCols("ForwardTrades") = FindColumnWithAliases(wsCand, 1, "ForwardTrades", "forward_trades")
    candCols("ForwardWin") = FindColumnWithAliases(wsCand, 1, "ForwardWin", "forward_winrate")
    candCols("WinCI_L") = FindColumnWithAliases(wsCand, 1, "WinCI_L", "forward_win_ci_low")
    candCols("WinCI_H") = FindColumnWithAliases(wsCand, 1, "WinCI_H", "forward_win_ci_high")
    candCols("ExpBootMean") = FindColumnWithAliases(wsCand, 1, "ExpBootMean", "forward_exp_boot_mean")
    candCols("ExpBootLow") = FindColumnWithAliases(wsCand, 1, "ExpBootLow", "forward_exp_boot_low")
    candCols("ExpBootHigh") = FindColumnWithAliases(wsCand, 1, "ExpBootHigh", "forward_exp_boot_high")
    candCols("ForwardAvgBars") = FindColumnWithAliases(wsCand, 1, "ForwardAvgBars", "forward_avg_bars")

    Dim r As Long
    For r = firstRow To lastRow
        Dim t As String: t = CStr(wsDash.Cells(r, dTickerCol).value)
        If Len(t) = 0 Then GoTo NextAttach
        If Not map.Exists(t) Then GoTo NextAttach
        cr = map(t)
        Dim k As Variant
        For Each k In dashCols.Keys
            Dim dc As Long: dc = dashCols(k)
            Dim cc As Long: cc = candCols(k)
            If dc > 0 And cc > 0 Then
                wsDash.Cells(r, dc).value = wsCand.Cells(cr, cc).value
            End If
        Next k
NextAttach:
    Next r
    On Error GoTo 0
End Sub

Private Function FindColumnWithAliases(ByVal ws As Worksheet, ByVal headerRow As Long, ParamArray aliases() As Variant) As Long
    ' Try all aliases in order; returns first match or 0
    Dim i As Long
    For i = LBound(aliases) To UBound(aliases)
        Dim name As String
        name = CStr(aliases(i))
        Dim c As Long
        c = FindColumn(ws, headerRow, name)
        If c > 0 Then
            FindColumnWithAliases = c
            Exit Function
        End If
    Next i
    FindColumnWithAliases = 0
End Function

Private Function ParseTime(ByVal value As Variant, ByVal fallback As String) As Date
    On Error GoTo Fail
    If IsDate(value) Then
        ParseTime = CDate(value)
        Exit Function
    End If
    If IsDate(fallback) Then
        ParseTime = CDate(fallback)
        Exit Function
    End If
Fail:
    ParseTime = TimeValue("09:00")
End Function

Private Function ComputeDynamicQty(ByVal basePrice As Double, ByVal side As String, ByVal maxBudget As Double, ByVal lotStep As Long, ByVal slipBp As Double, ByVal fallbackQty As Long) As Long
    Dim stepSize As Long: stepSize = IIf(lotStep > 0, lotStep, 1)
    Dim budget As Double: budget = IIf(maxBudget > 0, maxBudget, 0)
    Dim qty As Long
    Dim useQty As Long: useQty = IIf(fallbackQty > 0, fallbackQty, stepSize)
    If basePrice <= 0 Or budget <= 0 Then
        ComputeDynamicQty = useQty
        Exit Function
    End If
    Dim slipFactor As Double: slipFactor = Abs(slipBp) / 10000#
    Dim worstPrice As Double
    If UCase$(side) = "BUY" Then
        worstPrice = basePrice * (1 + slipFactor)
    Else
        worstPrice = basePrice * (1 - slipFactor)
    End If
    If worstPrice <= 0 Then worstPrice = basePrice
    qty = CLng(Fix(budget / worstPrice / stepSize)) * stepSize
    If qty <= 0 Then qty = stepSize
    ComputeDynamicQty = qty
End Function

Private Sub PlaceBracketIfAvailable(ByVal ws As Worksheet, ByVal r As Long, ByVal ticker As String, ByVal side As String, ByVal entryPx As Variant, ByVal qty As Long)
    On Error Resume Next
    Dim tpCol As Long: tpCol = FindColumn(ws, DASH_HEADER_ROW, "TPk")
    Dim slCol As Long: slCol = FindColumn(ws, DASH_HEADER_ROW, "SLk")
    Dim atrPriceCol As Long: atrPriceCol = FindColumn(ws, DASH_HEADER_ROW, "ATR")
    If tpCol = 0 Or slCol = 0 Or atrPriceCol = 0 Then Exit Sub
    Dim tpK As Double: tpK = CDbl(IfZero(ws.Cells(r, tpCol).value, 0))
    Dim slK As Double: slK = CDbl(IfZero(ws.Cells(r, slCol).value, 0))
    Dim atr As Double: atr = CDbl(IfZero(ws.Cells(r, atrPriceCol).value, 0))
    If atr <= 0 Then Exit Sub
    Dim tpPx As Double, slPx As Double
    If UCase$(side) = "BUY" Then
        tpPx = CDbl(entryPx) + tpK * atr
        slPx = CDbl(entryPx) - slK * atr
    Else
        tpPx = CDbl(entryPx) - tpK * atr
        slPx = CDbl(entryPx) + slK * atr
    End If
    PlaceOrder ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), tpPx, qty, "TP"
    PlaceOrder ticker, IIf(UCase$(side) = "BUY", "SELL", "BUY"), slPx, qty, "SL"
End Sub

Private Sub ScheduleCloseExit(ByVal ws As Worksheet, ByVal ticker As String, ByVal side As String)
    On Error Resume Next
    Dim t As Date
    t = ParseTime(ws.Range(DASH_CLOSE_TIME_CELL).value, DEFAULT_CLOSE_TIME)
    closeTimer = Date + t
    Application.OnTime EarliestTime:=closeTimer, Procedure:="AutoTrader.CloseAtMarket", Schedule:=True
    closeTimerScheduled = True
End Sub

Private Sub CancelCloseExit()
    On Error Resume Next
    If closeTimerScheduled And closeTimer <> 0 Then
        Application.OnTime EarliestTime:=closeTimer, Procedure:="AutoTrader.CloseAtMarket", Schedule:=False
    End If
    closeTimerScheduled = False
    closeTimer = 0
End Sub

Public Sub CloseAtMarket()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim tickerCol As Long: tickerCol = FindColumn(ws, DASH_HEADER_ROW, "Ticker")
    Dim priceCol As Long: priceCol = FindColumn(ws, DASH_HEADER_ROW, "PrevClose")
    If priceCol = 0 Then priceCol = FindColumn(ws, DASH_HEADER_ROW, "PreOpenMid")
    Dim qtyCol As Long: qtyCol = FindColumn(ws, DASH_HEADER_ROW, "DynamicQty")
    If tickerCol = 0 Or priceCol = 0 Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, tickerCol).End(xlUp).ROW
    Dim r As Long
    For r = DASH_DATA_START To lastRow
        Dim tkr As String: tkr = CStr(ws.Cells(r, tickerCol).value)
        If Len(tkr) > 0 Then
            Dim qty As Long
            qty = CLng(IfZero(IIf(qtyCol > 0, ws.Cells(r, qtyCol).value, 0), CLng(IfZero(ws.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))))
            PlaceOrder tkr, "FLAT", ws.Cells(r, priceCol).value, qty, "MOC"
        End If
    Next r
End Sub

Private Sub PlaceOrder(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    If qty <= 0 Then
        qty = CLng(IfZero(ws.Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))
    End If
    If CLng(IfZero(ws.Range(DASH_LIVE_CELL).value, 0)) = 1 Then
        PlaceOrderLive ticker, side, price, qty, info
    Else
        PlaceOrderDryRun ticker, side, price, qty, info
    End If
End Sub

Private Sub PlaceOrderLive(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_DASHBOARD)
    Dim macroName As String: macroName = CStr(IfZero(ws.Range(DASH_ORDERMACRO_CELL).value, ""))
    If Len(macroName) > 0 Then
        Application.Run macroName, ticker, side, price, qty, info
    Else
        ' fallback: dry run if macro not configured
        PlaceOrderDryRun ticker, side, price, qty, "NO_MACRO:" & info
    End If
End Sub

Private Function GetPrevJ(ByVal ticker As String) As Double
    If prevJ Is Nothing Then Set prevJ = CreateObject("Scripting.Dictionary")
    If prevJ.Exists(ticker) Then
        GetPrevJ = prevJ(ticker)
    Else
        GetPrevJ = 0
    End If
End Function

Private Sub SetPrevJ(ByVal ticker As String, ByVal value As Double)
    If prevJ Is Nothing Then Set prevJ = CreateObject("Scripting.Dictionary")
    prevJ(ticker) = value
End Sub

Private Sub PlaceOrderDryRun(ByVal ticker As String, ByVal side As String, ByVal price As Variant, ByVal qty As Long, ByVal info As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_ORDERS)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).ROW + 1
    ws.Cells(r, 1).value = Now
    ws.Cells(r, 2).value = ticker
    ws.Cells(r, 3).value = side
    ws.Cells(r, 4).value = price
    ws.Cells(r, 5).value = qty
    ws.Cells(r, 6).value = "DEMO " & info

    If qty <= 0 Then qty = CLng(IfZero(EnsureSheet(SHEET_DASHBOARD).Range(DASH_QTY_CELL).value, DEFAULT_ORDER_QTY))

    Dim wsPnl As Worksheet
    Set wsPnl = EnsureSheet("PnL")
    Dim pr As Long
    pr = wsPnl.Cells(wsPnl.Rows.Count, 1).End(xlUp).ROW + 1
    If pr < 5 Then pr = 5
    wsPnl.Cells(pr, 1).value = Now
    wsPnl.Cells(pr, 3).value = "DEMO"
    wsPnl.Cells(pr, 4).value = ticker
    wsPnl.Cells(pr, 5).value = side
    wsPnl.Cells(pr, 6).value = qty
    wsPnl.Cells(pr, 7).value = price
    wsPnl.Cells(pr, 9).value = info
End Sub

Private Function CheckRiskAndBudgetLimits(ByVal ticker As String, ByVal sessionKey As String, ByVal riskVal As Double, ByVal totalBudget As Double) As Boolean
    ' Enforce simple per-ticker and global caps using running usage dictionaries
    If riskVal <= 0 Or totalBudget <= 0 Then
        CheckRiskAndBudgetLimits = False
        Exit Function
    End If
    Dim usedTick As Double: usedTick = 0
    If Not tickerRiskUsed Is Nothing Then
        If tickerRiskUsed.Exists(ticker) Then usedTick = tickerRiskUsed(ticker)
    End If
    Dim capTick As Double
    capTick = totalBudget * RISK_PER_TICKER_FRAC
    Dim availTick As Double
    availTick = capTick - usedTick
    If availTick <= 0 Then
        CheckRiskAndBudgetLimits = False
        Exit Function
    End If

    Dim capTotal As Double
    capTotal = totalBudget * RISK_TOTAL_FRAC
    Dim availTotal As Double
    availTotal = capTotal - totalRiskUsed
    If availTotal <= 0 Then
        CheckRiskAndBudgetLimits = False
        Exit Function
    End If

    If riskVal > availTick Or riskVal > availTotal Then
        CheckRiskAndBudgetLimits = False
    Else
        CheckRiskAndBudgetLimits = True
    End If
End Function





