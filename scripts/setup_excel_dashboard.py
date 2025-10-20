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
    "Ticker", "\u9298\u67c4\u540d", "\u73fe\u5728\u306eJ\u5024", "\u95be\u5024\u4e56\u96e2\u7387(%)", "\u30b7\u30b0\u30ca\u30eb\u70b9\u706f", "\u30b7\u30b0\u30ca\u30eb\u7a2e\u5225", "\u73fe\u5728\u5024", "\u51fa\u6765\u9ad8\u52a0\u91cd\u5e73\u5747", "Selected", "SignalMode", "Session", "ATR_n", "TPk", "SLk", "J_th",
    "ForwardPF", "ForwardTrades", "WinCI_L", "WinCI_H", "ExpBootMean", "ExpBootLow", "ExpBootHigh",
    "ForwardAvgBars", "GapBucket", "GapRule", "GapSummary", "\u524d\u65e5\u7d42\u5024", "\u6c17\u914d\u5024(\u8cb7)", "\u6c17\u914d\u5024(\u58f2)", "\u6c17\u914d\u5024(\u4e2d\u592e)",
    "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7(bp)", "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7\u5e2f", "\u30e9\u30a4\u30d6\u30a2\u30af\u30b7\u30e7\u30f3", "DynamicQty"
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
    name_header = "銘柄名"
    current_j_header = "現在のJ値"
    gap_header = "閾値乖離率(%)"
    status_header = "シグナル点灯"
    kind_header = "シグナル種別"
    last_header = "現在値"
    vwap_header = "出来高加重平均"
    prev_header = "前日終値"
    bid_header = "気配値(買)"
    ask_header = "気配値(売)"
    mid_header = "気配値(中央)"
    gap_bp_header = "ライブギャップ(bp)"
    gap_bucket_header = "ライブギャップ帯"
    gap_action_header = "ライブアクション"

    for idx, header in enumerate(HEADERS):
        ws.Cells(5, HEADER_OFFSET_COL + idx).Value = header

    formula_rows = 400
    start_row = 6
    end_row = start_row + formula_rows
    ticker_pos = HEADERS.index("Ticker")

    def r1c(from_idx: int, to_idx: int) -> str:
        offset = from_idx - to_idx
        return "RC" if offset == 0 else f"RC[{offset}]"

    def set_formula(header: str, formula: str) -> None:
        if header not in HEADERS:
            return
        col_idx = HEADER_OFFSET_COL + HEADERS.index(header)
        rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
        try:
            rng.FormulaR1C1 = formula
        except Exception:
            pass

    if name_header in HEADERS:
        idx = HEADERS.index(name_header)
        ref = r1c(ticker_pos, idx)
        set_formula(name_header, f'=IF({ref}="","",IFERROR(@RssMarket({ref},"銘柄名"),""))')

    if current_j_header in HEADERS:
        idx = HEADERS.index(current_j_header)
        last_idx = HEADERS.index(last_header) if last_header in HEADERS else None
        vwap_idx = HEADERS.index(vwap_header) if vwap_header in HEADERS else None
        atr_idx = HEADERS.index("ATR_n") if "ATR_n" in HEADERS else None
        if last_idx is not None and vwap_idx is not None and atr_idx is not None:
            last_ref = r1c(last_idx, idx)
            vwap_ref = r1c(vwap_idx, idx)
            atr_ref = r1c(atr_idx, idx)
            set_formula(current_j_header, f'=IF(OR({last_ref}="",{vwap_ref}="",{atr_ref}=0),"",({last_ref}-{vwap_ref})/{atr_ref})')
        else:
            set_formula(current_j_header, '=""')

    if gap_header in HEADERS:
        idx = HEADERS.index(gap_header)
        current_idx = HEADERS.index(current_j_header) if current_j_header in HEADERS else None
        jth_idx = HEADERS.index("J_th") if "J_th" in HEADERS else None
        if current_idx is not None and jth_idx is not None:
            current_ref = r1c(current_idx, idx)
            jth_ref = r1c(jth_idx, idx)
            set_formula(gap_header, f'=IF(OR({jth_ref}="",{current_ref}="",{jth_ref}=0),"",MAX(0,(ABS({jth_ref})-ABS({current_ref}))/ABS({jth_ref})*100))')
            rng_gap = ws.Range(ws.Cells(start_row, HEADER_OFFSET_COL + idx), ws.Cells(end_row, HEADER_OFFSET_COL + idx))
            try:
                rng_gap.FormatConditions.Delete()
                cf = rng_gap.FormatConditions.AddColorScale(3)
                cf.ColorScaleCriteria(1).Type = 0
                cf.ColorScaleCriteria(1).Value = 50
                cf.ColorScaleCriteria(1).FormatColor.Color = 0xC0504D
                cf.ColorScaleCriteria(2).Type = 0
                cf.ColorScaleCriteria(2).Value = 25
                cf.ColorScaleCriteria(2).FormatColor.Color = 0x92D050
                cf.ColorScaleCriteria(3).Type = 0
                cf.ColorScaleCriteria(3).Value = 0
                cf.ColorScaleCriteria(3).FormatColor.Color = 0x548235
            except Exception:
                pass
        else:
            set_formula(gap_header, '=""')

    if status_header in HEADERS:
        set_formula(status_header, '=""')
    if kind_header in HEADERS:
        set_formula(kind_header, '=""')

    if last_header in HEADERS:
        idx = HEADERS.index(last_header)
        ref = r1c(ticker_pos, idx)
        set_formula(last_header, f'=IF({ref}="","",IFERROR(@RssMarket({ref},"現在値"),""))')

    if vwap_header in HEADERS:
        idx = HEADERS.index(vwap_header)
        ref = r1c(ticker_pos, idx)
        set_formula(vwap_header, f'=IF({ref}="","",IFERROR(@RssMarket({ref},"出来高加重平均"),""))')

    formula_map = [
        (prev_header, 15),
        (bid_header, 56),
        (ask_header, 55),
    ]
    for header, code in formula_map:
        if header in HEADERS:
            idx = HEADERS.index(header)
            ref = r1c(ticker_pos, idx)
            set_formula(header, f'=IF({ref}="","",IFERROR(@RssMarket({ref},{code}),""))')

    if mid_header in HEADERS and bid_header in HEADERS and ask_header in HEADERS:
        mid_idx = HEADERS.index(mid_header)
        bid_idx = HEADERS.index(bid_header)
        ask_idx = HEADERS.index(ask_header)
        bid_ref = r1c(bid_idx, mid_idx)
        ask_ref = r1c(ask_idx, mid_idx)
        set_formula(mid_header, f'=IF(OR({bid_ref}="",{ask_ref}=""),"",({bid_ref}+{ask_ref})/2)')

    if gap_bp_header in HEADERS and mid_header in HEADERS and prev_header in HEADERS:
        gap_idx = HEADERS.index(gap_bp_header)
        mid_idx = HEADERS.index(mid_header)
        prev_idx = HEADERS.index(prev_header)
        mid_ref = r1c(mid_idx, gap_idx)
        prev_ref = r1c(prev_idx, gap_idx)
        set_formula(gap_bp_header, f'=IF(OR({mid_ref}="",{prev_ref}=""),"",({mid_ref}-{prev_ref})/{prev_ref}*10000)')

    if gap_bucket_header in HEADERS and gap_bp_header in HEADERS:
        bucket_idx = HEADERS.index(gap_bucket_header)
        gap_idx = HEADERS.index(gap_bp_header)
        gap_ref = r1c(gap_idx, bucket_idx)
        set_formula(gap_bucket_header, f'=IF({gap_ref}="","",IF(ABS({gap_ref})>=120,">=120bp",IF(ABS({gap_ref})>=80,"80-120bp",IF(ABS({gap_ref})>=50,"50-80bp","<50bp"))))')

    if gap_action_header in HEADERS and gap_bucket_header in HEADERS:
        action_idx = HEADERS.index(gap_action_header)
        bucket_idx = HEADERS.index(gap_bucket_header)
        bucket_ref = r1c(bucket_idx, action_idx)
        set_formula(gap_action_header, f'=IF({bucket_ref}="","",IF({bucket_ref}=">=120bp","j-cross only; TP-0.2; SL+0.2",IF({bucket_ref}="80-120bp","Skip opposite; J_th+0.2",IF({bucket_ref}="50-80bp","J_th+0.1","Baseline")))))')

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


