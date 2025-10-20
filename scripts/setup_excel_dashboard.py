import win32com.client
from pathlib import Path

# Install buttons on NewDashboard to avoid conflicts with existing Dashboard
WB_PATH = Path('C:/AI/asagake/SHINSOKU.xlsm')
SHEET_NAME = "NewDashboard"
BUTTONS = [
    ("btnLoadCandidates", "Load Candidates", "AutoTrader.ButtonLoadCandidates", 1),
    ("btnPushDashboard", "Push Candidates", "AutoTrader.ButtonPushCandidates", 2),
    ("btnStartAuto", "Start Auto", "AutoTrader.ButtonStartAuto", 3),
    ("btnStopAuto", "Stop Auto", "AutoTrader.ButtonStopAuto", 4),
    ("btnRefreshAuto", "Refresh Now", "AutoTrader.ButtonRefreshNow", 5),
    ("btnCatchUp", "Catch Up (Nightly)", "AutoTrader.ButtonCatchUp", 6),
]

HEADER_OFFSET_COL = 8
HEADERS = [
    "Ticker", "Selected", "SignalMode", "Session", "ATR_n", "TPk", "SLk", "J_th",
    "ForwardPF", "ForwardTrades", "WinCI_L", "WinCI_H", "ExpBootMean",
    "ExpBootLow", "ExpBootHigh", "ForwardAvgBars", "GapBucket", "GapRule",
    "GapSummary", "PrevClose", "PreOpenBid", "PreOpenAsk", "PreOpenMid",
    "LiveGapBp", "LiveGapBucket", "LiveGapAction", "DynamicQty"
]

CONFIG_CELLS = {
    "A2": "AutoTrade Status (0=Off,1=On)",
    "B2": 0,
    "A3": "Daily Max Orders",
    "B3": 20,
    "A4": "Session Start (HH:MM)",
    "B4": "09:00",
    "A5": "Session End (HH:MM)",
    "B5": "09:15"
}

if not WB_PATH.exists():
    raise SystemExit(f"Workbook not found: {WB_PATH}")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
try:
    wb = excel.Workbooks.Open(str(WB_PATH))
    # 既存の補助シートは一度削除
    for name in ("設定説明", "ボタン説明", "運用ガイド", "SettingsGuide", "ButtonGuide", "Operations"):
        try:
            wb.Worksheets(name).Delete()
        except Exception:
            pass
    # Ensure NewDashboard exists
    try:
        ws = wb.Worksheets(SHEET_NAME)
    except Exception:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_NAME

    # Apply config labels/defaults
    for addr, value in CONFIG_CELLS.items():
        cell = ws.Range(addr)
        if isinstance(value, str):
            if cell.Value in (None, ""):
                cell.Value = value
            else:
                # Keep user overrides for string values except labels (column A)
                if addr.startswith("A"):
                    cell.Value = value
        else:
            if cell.Value in (None, ""):
                cell.Value = value

    # Header extensions
    for idx, header in enumerate(HEADERS):
        ws.Cells(5, HEADER_OFFSET_COL + idx).Value = header

    formula_rows = 400
    start_row = 6
    end_row = start_row + formula_rows
    ticker_pos = HEADERS.index("Ticker")
    formula_map = {
        "PrevClose": '=IFERROR(RssMarket({ticker},15),"")',
        "PreOpenBid": '=IFERROR(RssMarket({ticker},56),"")',
        "PreOpenAsk": '=IFERROR(RssMarket({ticker},55),"")',
        "PreOpenMid": '=IF(OR(RC[-1]="",RC[-2]=""),"", (RC[-1]+RC[-2])/2)',
        "LiveGapBp": '=IF(OR(RC[-1]="",RC[-4]=""),"", (RC[-1]-RC[-4])/RC[-4]*10000)',
        "LiveGapBucket": '=IF(RC[-1]="","",IF(ABS(RC[-1])>=120,">=120bp",IF(ABS(RC[-1])>=80,"80-120bp",IF(ABS(RC[-1])>=50,"50-80bp","<50bp"))))',
        "LiveGapAction": '=IF(RC[-1]="","",IF(RC[-1]=">=120bp","j-cross only; TP-0.2; SL+0.2",IF(RC[-1]="80-120bp","Skip opposite; J_th+0.2",IF(RC[-1]="50-80bp","J_th+0.1","Baseline"))))',
    }
    for header, formula in formula_map.items():
        if header in HEADERS:
            col_idx = HEADER_OFFSET_COL + HEADERS.index(header)
            ticker_offset = ticker_pos - HEADERS.index(header)
            if ticker_offset == 0:
                ticker_ref = "RC"
            else:
                ticker_ref = f"RC[{ticker_offset}]"
            rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
            rng.FormulaR1C1 = formula.format(ticker=ticker_ref)

    # 設定セルのラベルと初期値
    labels = [
        ("A2", "AutoTrade (0/1)", "B2", 0),
        ("A3", "Daily Max Orders", "B3", 20),
        ("A4", "Session Start", "B4", "09:00"),
        ("A5", "Session End", "B5", "09:15"),
        ("A6", "初期選択フラグ (0/1)", "B6", 1),
        ("A7", "再エントリー許可 (0/1)", "B7", 0),
        ("A8", "非常停止 (0/1)", "B8", 0),
        ("A9", "実発注 (0/1)", "B9", 0),
        ("A10", "発注マクロ名", "B10", "MS2Bridge.Place"),
        ("A11", "引け成行時刻 (HH:MM:SS)", "B11", "14:59:30"),
        ("A12", "注文数量の既定値", "B12", 100),
        ("A13", "注文種別 (TIF/Type)", "B13", "MKT"),
        ("A14", "1回あたり予算上限 (JPY)", "B14", 10000000),
        ("A15", "ロット刻み (株)", "B15", 100),
        ("A16", "許容スリッページ (bp)", "B16", 30),
    ]
    for label_cell, label_text, value_cell, default in labels:
        ws.Range(label_cell).Value = label_text
        if ws.Range(value_cell).Value in (None, ""):
            ws.Range(value_cell).Value = default

    # 既存の Candidates / Orders / MS2_Config を確保
    for sheet_name in ("Candidates", "Orders", "MS2_Config"):
        try:
            wb.Worksheets(sheet_name)
        except Exception:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = sheet_name
    ws_orders = wb.Worksheets("Orders")
    if ws_orders.Cells(1, 1).Value in (None, ""):
        ws_orders.Range("A1:F1").Value = ("Time", "Ticker", "Side", "Price", "Qty", "Note")

    # Seed MS2_Config with extended parameter columns

    ws_cfg = wb.Worksheets("MS2_Config")

    if ws_cfg.Cells(1, 1).Value in (None, ""):

        headers = [

            "Key",

            "Value",

            "Function",

            "OrderIdFmt",

            "BuyCode",

            "SellCode",

            "OrderDiv",

            "SorDiv",

            "CreditDiv",

            "PriceDiv",

            "ExecCond",

            "Term",

            "AccountDiv",

            "TriggerPrice1",

            "TriggerCond1",

            "TriggerPrice2",

            "TriggerCond2",

            "SetDiv",

            "SetPriceDiv",

            "SetPrice",

            "SetExecCond",

            "SetAccount",

            "DefaultPrice",

            "Notes",

        ]

        ws_cfg.Range(ws_cfg.Cells(1, 1), ws_cfg.Cells(1, len(headers))).Value = (headers,)

        default_rows = [

            {

                "Key": "Account",

                "Value": "",

                "Notes": "RSS account code (fill manually).",

            },

            {

                "Key": "Market",

                "Value": "TSE",

                "Notes": "Default market code (e.g. TSE or JNX).",

            },

            {

                "Key": "EntryTemplate",

                "Value": "RssStockOrder_v({OrderId},{TickerCode},{SideCode},{OrderDiv},{SorDiv},{Qty},{PriceDiv},{OrderPrice},{ExecCond},{Term},{AccountDiv},{TriggerPrice1},{TriggerCond1},{TriggerPrice2},{TriggerCond2},{SetDiv},{SetPriceDiv},{SetPrice},{SetExecCond},{SetAccount})",

                "Function": "RssStockOrder_v",

                "OrderIdFmt": "{Ticker}-{Info}-{Time}",

                "BuyCode": 1,

                "SellCode": 2,

                "OrderDiv": 1,

                "SorDiv": 0,

                "CreditDiv": "",

                "PriceDiv": 0,

                "ExecCond": 0,

                "Term": 0,

                "AccountDiv": "",

                "TriggerPrice1": "",

                "TriggerCond1": 0,

                "TriggerPrice2": "",

                "TriggerCond2": 0,

                "SetDiv": 0,

                "SetPriceDiv": 0,

                "SetPrice": "",

                "SetExecCond": 0,

                "SetAccount": 0,

                "DefaultPrice": "",

                "Notes": "Entry template for market orders; adjust placeholders for broker specs.",

            },

            {

                "Key": "TPTemplate",

                "Value": "RssStockOrder_v({OrderId},{TickerCode},{SideCode},{OrderDiv},{SorDiv},{Qty},{PriceDiv},{OrderPrice},{ExecCond},{Term},{AccountDiv},{TriggerPrice1},{TriggerCond1},{TriggerPrice2},{TriggerCond2},{SetDiv},{SetPriceDiv},{SetPrice},{SetExecCond},{SetAccount})",

                "Function": "RssStockOrder_v",

                "OrderIdFmt": "{Ticker}-TP-{Time}",

                "BuyCode": 2,

                "SellCode": 1,

                "OrderDiv": 2,

                "SorDiv": 0,

                "CreditDiv": "",

                "PriceDiv": 1,

                "ExecCond": 0,

                "Term": 0,

                "AccountDiv": "",

                "TriggerPrice1": "",

                "TriggerCond1": 0,

                "TriggerPrice2": "",

                "TriggerCond2": 0,

                "SetDiv": 0,

                "SetPriceDiv": 0,

                "SetPrice": "",

                "SetExecCond": 0,

                "SetAccount": 0,

                "DefaultPrice": "",

                "Notes": "Take-profit template; make sure SideCode flips relative to entry.",

            },

            {

                "Key": "SLTemplate",

                "Value": "RssStockOrder_v({OrderId},{TickerCode},{SideCode},{OrderDiv},{SorDiv},{Qty},{PriceDiv},{OrderPrice},{ExecCond},{Term},{AccountDiv},{TriggerPrice1},{TriggerCond1},{TriggerPrice2},{TriggerCond2},{SetDiv},{SetPriceDiv},{SetPrice},{SetExecCond},{SetAccount})",

                "Function": "RssStockOrder_v",

                "OrderIdFmt": "{Ticker}-SL-{Time}",

                "BuyCode": 2,

                "SellCode": 1,

                "OrderDiv": 2,

                "SorDiv": 0,

                "CreditDiv": "",

                "PriceDiv": 1,

                "ExecCond": 0,

                "Term": 0,

                "AccountDiv": "",

                "TriggerPrice1": "",

                "TriggerCond1": 0,

                "TriggerPrice2": "",

                "TriggerCond2": 0,

                "SetDiv": 0,

                "SetPriceDiv": 0,

                "SetPrice": "",

                "SetExecCond": 0,

                "SetAccount": 0,

                "DefaultPrice": "",

                "Notes": "Stop-loss template; configure triggers when sending stop orders.",

            },

            {

                "Key": "MOCTemplate",

                "Value": "RssStockOrder_v({OrderId},{TickerCode},{SideCode},{OrderDiv},{SorDiv},{Qty},{PriceDiv},{OrderPrice},{ExecCond},{Term},{AccountDiv},{TriggerPrice1},{TriggerCond1},{TriggerPrice2},{TriggerCond2},{SetDiv},{SetPriceDiv},{SetPrice},{SetExecCond},{SetAccount})",

                "Function": "RssStockOrder_v",

                "OrderIdFmt": "{Ticker}-MOC-{Date}",

                "BuyCode": 2,

                "SellCode": 1,

                "OrderDiv": 1,

                "SorDiv": 0,

                "CreditDiv": "",

                "PriceDiv": 0,

                "ExecCond": 0,

                "Term": 0,

                "AccountDiv": "",

                "TriggerPrice1": "",

                "TriggerCond1": 0,

                "TriggerPrice2": "",

                "TriggerCond2": 0,

                "SetDiv": 0,

                "SetPriceDiv": 0,

                "SetPrice": "",

                "SetExecCond": 0,

                "SetAccount": 0,

                "DefaultPrice": "",

                "Notes": "Market-on-close template; adjust order type if broker requires.",

            },

        ]

        start_row = 2

        for idx, row in enumerate(default_rows):

            r = start_row + idx

            for col_idx, header in enumerate(headers, start=1):

                if header in row:

                    ws_cfg.Cells(r, col_idx).Value = row[header]

        ws_cfg.Columns.AutoFit()

# Remove existing buttons by name
    for btn_name, _, _, _ in BUTTONS:
        for shp in list(ws.Shapes):
            if shp.Name == btn_name:
                shp.Delete()

    # Add buttons stacked below header row (row 1 area)
    left_start = ws.Cells(2, 4).Left
    top_start = ws.Cells(2, 4).Top
    width = 120
    height = 24
    spacing = 6
    for idx, (btn_name, caption, macro, order) in enumerate(BUTTONS):
        btn = ws.Buttons().Add(left_start, top_start + (height + spacing) * idx, width, height)
        btn.Name = btn_name
        btn.OnAction = macro
        btn.Text = caption

    
        # 補助シートを作成（日本語）
    def ensure_sheet(name: str):
        try:
            return wb.Worksheets(name)
        except Exception:
            sht = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            sht.Name = name
            return sht

    ws_help = ensure_sheet("設定説明")
    ws_help.Cells.Clear()
    ws_help.Range("A1").Value = "NewDashboard 設定一覧"
    guide_rows = [
        ("B2", "AutoTrade (0/1): 1 で自動監視を開始"),
        ("B3", "Daily Max Orders: 1 日の最大発注回数"),
        ("B4", "Session Start: シグナル判定の開始時刻"),
        ("B5", "Session End: シグナル判定の終了時刻"),
        ("B6", "初期選択フラグ: Push 後に空欄ならこの値をセット"),
        ("B7", "再エントリー許可: 1=毎朝 Selected を既定値に戻す / 0=前日の設定を保持"),
        ("B8", "非常停止: 1 で新規エントリーを即停止"),
        ("B9", "実発注: 0=ドライラン, 1=MS2Bridge 経由で発注"),
        ("B10", "発注マクロ名: 実発注時に呼び出すマクロ"),
        ("B11", "引け成行時刻: CloseAtMarket で使用"),
        ("B12", "注文数量の既定値: 動的ロットが計算できない場合に使用"),
        ("B13", "注文種別 (TIF/Type): エントリーの注文タイプ"),
        ("B14", "1回あたり予算上限 (JPY): 動的ロット算出に使用"),
        ("B15", "ロット刻み (株): 数量をこの刻みに丸める"),
        ("B16", "許容スリッページ (bp): 予算に対する余裕設定"),
    ]
    for addr, text_help in guide_rows:
        ws_help.Range(addr).Value = text_help
    ws_help.Columns.AutoFit()

    ws_buttons = ensure_sheet("ボタン説明")
    ws_buttons.Cells.Clear()
    ws_buttons.Range("A1").Value = "ボタン一覧"
    button_lines = [
        "Load Candidates: 夜間バッチ出力 (output/excel/candidates_nextday.csv) を Candidates に読込", 
        "Push Candidates: Candidates の内容を NewDashboard に展開", 
        "Start Auto: 1 秒間隔の監視ループ開始", 
        "Stop Auto: 監視ループ停止", 
        "Refresh Now: その場で 1 回だけ評価を実行", 
        "Catch Up (Nightly): 夜間バッチスクリプトを手動実行" 
    ]
    for idx, line in enumerate(button_lines, start=2):
        ws_buttons.Range(f"A{idx}").Value = line
    ws_buttons.Columns.AutoFit()

    ws_ops = ensure_sheet("運用ガイド")
    ws_ops.Cells.Clear()
    ws_ops.Range("A1").Value = "1日の流れ"
    ops_lines = [
        "18:00頃 Nightly 実行 (Yahoo 売買代金上位300 → coarse/refine → Gap 集計)",
        "朝: SHINSOKU.xlsm を開き Load → Push で候補を反映", 
        "候補の Selected / GapRule / DynamicQty を確認・調整", 
        "AutoTrade=1, Live=0/1 を確認し Start Auto", 
        "DynamicQty = floor(予算 ÷ (価格×(1+slip))) を刻みに丸めて使用", 
        "TP/SL はテンプレートに従って指値、Close-Out Time で引け成行", 
        "ドライランは Orders/PnL に DEMO として記録" 
    ]
    for idx, line in enumerate(ops_lines, start=2):
        ws_ops.Range(f"A{idx}").Value = line
    ws_ops.Columns.AutoFit()

    wb.Save()
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()


