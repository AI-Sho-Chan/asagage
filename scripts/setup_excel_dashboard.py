import contextlib
from pathlib import Path

import win32com.client


WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
SHEET_NAME = "NewDashboard"

BUTTONS = [
    ("btnLoadCandidates", "Load Candidates", "AutoTrader.ButtonLoadCandidates", 1),
    ("btnPushDashboard", "Push Candidates", "AutoTrader.ButtonPushCandidates", 2),
    ("btnStartAuto", "Start Auto", "AutoTrader.ButtonStartAuto", 3),
    ("btnStopAuto", "Stop Auto", "AutoTrader.ButtonStopAuto", 4),
    ("btnRefreshAuto", "Refresh Now", "AutoTrader.ButtonRefreshNow", 5),
    ("btnCatchUp", "Catch Up (Nightly)", "AutoTrader.ButtonCatchUp", 6),
]

HEADER_OFFSET_COL = 8  # column H
HEADER_ROW = 5
DATA_START_ROW = 6
FORMULA_ROWS = 400

HEADERS = [
    "Ticker",
    "\u9298\u67c4\u540d",
    "\u73fe\u5728\u306eJ\u5024",
    "\u95be\u5024\u4e56\u96e2\u7387(%)",
    "\u30b7\u30b0\u30ca\u30eb\u70b9\u706f",
    "\u30b7\u30b0\u30ca\u30eb\u7a2e\u5225",
    "\u73fe\u5728\u5024",
    "\u51fa\u6765\u9ad8\u52a0\u91cd\u5e73\u5747",
    "Selected",
    "SignalMode",
    "Session",
    "ATR_n",
    "TPk",
    "SLk",
    "J_th",
    "ForwardPF",
    "ForwardTrades",
    "ForwardWin",
    "WinCI_L",
    "WinCI_H",
    "ExpBootMean",
    "ExpBootLow",
    "ExpBootHigh",
    "ForwardAvgBars",
    "GapBucket",
    "GapRule",
    "GapSummary",
    "\u524d\u65e5\u7d42\u5024",
    "\u6c17\u914d\u5024(\u8cb7)",
    "\u6c17\u914d\u5024(\u58f2)",
    "\u6c17\u914d\u5024(\u4e2d\u592e)",
    "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7(bp)",
    "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7\u5e2f",
    "\u30e9\u30a4\u30d6\u30a2\u30af\u30b7\u30e7\u30f3",
    "DynamicQty",
]

CONFIG_CELLS = {
    "A2": "AutoTrade Status (0=Off,1=On)",
    "B2": 0,
    "A3": "Daily Max Orders",
    "B3": 20,
    "A4": "Session Start (HH:MM)",
    "B4": "09:00",
    "A5": "Session End (HH:MM)",
    "B5": "09:15",
}

SETTING_LABELS = [
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

NAME_FIELD = "\u9298\u67c4\u540d\u79f0"
LAST_FIELD = "\u73fe\u5728\u5024"
VWAP_FIELD = "\u51fa\u6765\u9ad8\u52a0\u91cd\u5e73\u5747"


def rgb(red: int, green: int, blue: int) -> int:
    return red | (green << 8) | (blue << 16)


def ensure_excel():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    try:
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except Exception:
        pass
    return excel


def open_workbook(excel):
    wb = excel.Workbooks.Open(str(WB_PATH))
    if wb is None:
        raise RuntimeError(f"Failed to open workbook: {WB_PATH}")
    return wb


def ensure_sheet(wb, name: str):
    try:
        return wb.Worksheets(name)
    except Exception:
        sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        sheet.Name = name
        return sheet


def clear_helper_sheets(wb):
    helper_names = ("設定説明", "ボタン説明", "運用ガイド", "SettingsGuide", "ButtonGuide", "Operations")
    for name in helper_names:
        with contextlib.suppress(Exception):
            wb.Worksheets(name).Delete()


def apply_config_defaults(ws):
    for addr, value in CONFIG_CELLS.items():
        cell = ws.Range(addr)
        if isinstance(value, str):
            if cell.Value in (None, "") or addr.startswith("A"):
                cell.Value = value
        else:
            if cell.Value in (None, ""):
                cell.Value = value


def write_headers(ws):
    for offset, header in enumerate(HEADERS):
        ws.Cells(HEADER_ROW, HEADER_OFFSET_COL + offset).Value = header


def column_index(header: str):
    if header not in HEADERS:
        return None
    return HEADER_OFFSET_COL + HEADERS.index(header)


def r1c(from_idx: int, to_idx: int) -> str:
    offset = from_idx - to_idx
    return "RC" if offset == 0 else f"RC[{offset}]"


def set_formula(ws, header: str, formula: str):
    col = column_index(header)
    if col is None:
        return
    end_row = DATA_START_ROW + FORMULA_ROWS
    rng = ws.Range(ws.Cells(DATA_START_ROW, col), ws.Cells(end_row, col))
    rng.FormulaR1C1 = formula
    return rng


def apply_realtime_formulas(ws):
    ticker_idx = HEADERS.index("Ticker")

    set_formula(ws, "\u30b7\u30b0\u30ca\u30eb\u70b9\u706f", '=""')
    set_formula(ws, "\u30b7\u30b0\u30ca\u30eb\u7a2e\u5225", '=""')

    name_idx = HEADERS.index("\u9298\u67c4\u540d")
    name_ref = r1c(ticker_idx, name_idx)
    set_formula(
        ws,
        "\u9298\u67c4\u540d",
        f'=IF({name_ref}="","",IFERROR(@RssMarket({name_ref},"{NAME_FIELD}"),""))',
    )

    last_idx = HEADERS.index("\u73fe\u5728\u5024")
    last_ref = r1c(ticker_idx, last_idx)
    set_formula(
        ws,
        "\u73fe\u5728\u5024",
        f'=IF({last_ref}="","",IFERROR(@RssMarket({last_ref},"{LAST_FIELD}"),""))',
    )

    vwap_idx = HEADERS.index("\u51fa\u6765\u9ad8\u52a0\u91cd\u5e73\u5747")
    vwap_ref = r1c(ticker_idx, vwap_idx)
    set_formula(
        ws,
        "\u51fa\u6765\u9ad8\u52a0\u91cd\u5e73\u5747",
        f'=IF({vwap_ref}="","",IFERROR(@RssMarket({vwap_ref},"{VWAP_FIELD}"),""))',
    )

    current_j_idx = HEADERS.index("\u73fe\u5728\u306eJ\u5024")
    atr_idx = HEADERS.index("ATR_n")
    last_ref_j = r1c(last_idx, current_j_idx)
    vwap_ref_j = r1c(vwap_idx, current_j_idx)
    atr_ref_j = r1c(atr_idx, current_j_idx)
    set_formula(
        ws,
        "\u73fe\u5728\u306eJ\u5024",
        f'=IF(OR({last_ref_j}="",{vwap_ref_j}="",{atr_ref_j}=0),"",({last_ref_j}-{vwap_ref_j})/{atr_ref_j})',
    )

    j_th_idx = HEADERS.index("J_th")
    gap_idx = HEADERS.index("\u95be\u5024\u4e56\u96e2\u7387(%)")
    current_ref_gap = r1c(current_j_idx, gap_idx)
    jth_ref = r1c(j_th_idx, gap_idx)
    gap_range = set_formula(
        ws,
        "\u95be\u5024\u4e56\u96e2\u7387(%)",
        f'=IF(OR({jth_ref}="",{current_ref_gap}="",{jth_ref}=0),"",MAX(0,(ABS({jth_ref})-ABS({current_ref_gap}))/ABS({jth_ref})*100))',
    )
    if gap_range is not None:
        with contextlib.suppress(Exception):
            gap_range.FormatConditions.Delete()
            cf = gap_range.FormatConditions.AddColorScale(3)
            cf.ColorScaleCriteria(1).Type = 0
            cf.ColorScaleCriteria(1).Value = 50
            cf.ColorScaleCriteria(1).FormatColor.Color = rgb(0, 112, 192)
            cf.ColorScaleCriteria(2).Type = 0
            cf.ColorScaleCriteria(2).Value = 25
            cf.ColorScaleCriteria(2).FormatColor.Color = rgb(146, 208, 80)
            cf.ColorScaleCriteria(3).Type = 0
            cf.ColorScaleCriteria(3).Value = 0
            cf.ColorScaleCriteria(3).FormatColor.Color = rgb(0, 176, 80)

    prev_idx = HEADERS.index("\u524d\u65e5\u7d42\u5024")
    prev_ref = r1c(ticker_idx, prev_idx)
    set_formula(ws, "\u524d\u65e5\u7d42\u5024", f'=IF({prev_ref}="","",IFERROR(@RssMarket({prev_ref},15),""))')

    bid_idx = HEADERS.index("\u6c17\u914d\u5024(\u8cb7)")
    bid_ref = r1c(ticker_idx, bid_idx)
    set_formula(ws, "\u6c17\u914d\u5024(\u8cb7)", f'=IF({bid_ref}="","",IFERROR(@RssMarket({bid_ref},56),""))')

    ask_idx = HEADERS.index("\u6c17\u914d\u5024(\u58f2)")
    ask_ref = r1c(ticker_idx, ask_idx)
    set_formula(ws, "\u6c17\u914d\u5024(\u58f2)", f'=IF({ask_ref}="","",IFERROR(@RssMarket({ask_ref},55),""))')

    mid_idx = HEADERS.index("\u6c17\u914d\u5024(\u4e2d\u592e)")
    bid_ref_mid = r1c(bid_idx, mid_idx)
    ask_ref_mid = r1c(ask_idx, mid_idx)
    set_formula(
        ws,
        "\u6c17\u914d\u5024(\u4e2d\u592e)",
        f'=IF(OR({bid_ref_mid}="",{ask_ref_mid}=""),"",({bid_ref_mid}+{ask_ref_mid})/2)',
    )

    gap_bp_idx = HEADERS.index("\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7(bp)")
    mid_ref_gap = r1c(mid_idx, gap_bp_idx)
    prev_ref_gap = r1c(prev_idx, gap_bp_idx)
    set_formula(
        ws,
        "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7(bp)",
        f'=IF(OR({mid_ref_gap}="",{prev_ref_gap}=""),"",({mid_ref_gap}-{prev_ref_gap})/{prev_ref_gap}*10000)',
    )

    bucket_idx = HEADERS.index("\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7\u5e2f")
    gap_ref_bucket = r1c(gap_bp_idx, bucket_idx)
    set_formula(
        ws,
        "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7\u5e2f",
        f'=IF({gap_ref_bucket}="","",IF(ABS({gap_ref_bucket})>=120,">=120bp",IF(ABS({gap_ref_bucket})>=80,"80-120bp",IF(ABS({gap_ref_bucket})>=50,"50-80bp","<50bp"))))',
    )

    action_idx = HEADERS.index("\u30e9\u30a4\u30d6\u30a2\u30af\u30b7\u30e7\u30f3")
    bucket_ref_action = r1c(bucket_idx, action_idx)
    set_formula(
        ws,
        "\u30e9\u30a4\u30d6\u30a2\u30af\u30b7\u30e7\u30f3",
        f'=IF({bucket_ref_action}="","",IF({bucket_ref_action}=">=120bp","j-cross only; TP-0.2; SL+0.2",IF({bucket_ref_action}="80-120bp","Skip opposite; J_th+0.2",IF({bucket_ref_action}="50-80bp","J_th+0.1","Baseline"))))',
    )


def apply_config_labels(ws):
    for label_cell, label_text, value_cell, default in SETTING_LABELS:
        ws.Range(label_cell).Value = label_text
        if ws.Range(value_cell).Value in (None, ""):
            ws.Range(value_cell).Value = default


def ensure_candidates_sheet(wb):
    ensure_sheet(wb, "Candidates")


def ensure_orders_sheet(wb):
    ws_orders = ensure_sheet(wb, "Orders")
    if ws_orders.Cells(1, 1).Value in (None, ""):
        ws_orders.Range("A1:F1").Value = ("Time", "Ticker", "Side", "Price", "Qty", "Note")


def seed_ms2_config(wb):
    ws_cfg = ensure_sheet(wb, "MS2_Config")
    if ws_cfg.Cells(1, 1).Value not in (None, ""):
        return

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
        target_row = start_row + idx
        for col_idx, header in enumerate(headers, start=1):
            if header in row:
                ws_cfg.Cells(target_row, col_idx).Value = row[header]

    ws_cfg.Columns.AutoFit()


def install_buttons(ws):
    for name, _, _, _ in BUTTONS:
        for shape in list(ws.Shapes):
            if shape.Name == name:
                shape.Delete()

    left = ws.Cells(2, 4).Left
    top = ws.Cells(2, 4).Top
    width = 120
    height = 24
    spacing = 6
    for idx, (name, caption, macro, _) in enumerate(BUTTONS):
        button = ws.Buttons().Add(left, top + idx * (height + spacing), width, height)
        button.Name = name
        button.OnAction = macro
        button.Text = caption


def ensure_support_sheets(wb):
    def ensure(name: str):
        return ensure_sheet(wb, name)

    ws_help = ensure("設定説明")
    ws_help.Cells.Clear()
    ws_help.Range("A1").Value = "NewDashboard 設定一覧"
    help_rows = [
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
    for addr, text in help_rows:
        ws_help.Range(addr).Value = text
    ws_help.Columns.AutoFit()

    ws_buttons = ensure("ボタン説明")
    ws_buttons.Cells.Clear()
    ws_buttons.Range("A1").Value = "ボタン一覧"
    button_lines = [
        "Load Candidates: output/excel/candidates_nextday.csv を Candidates に読み込み",
        "Push Candidates: Candidates の内容を NewDashboard に展開",
        "Start Auto: 1 秒間隔の監視ループ開始",
        "Stop Auto: 監視ループ停止",
        "Refresh Now: その場で 1 回だけ評価を実行",
        "Catch Up (Nightly): 夜間バッチスクリプトを手動実行",
    ]
    for idx, line in enumerate(button_lines, start=2):
        ws_buttons.Range(f"A{idx}").Value = line
    ws_buttons.Columns.AutoFit()

    ws_ops = ensure("運用ガイド")
    ws_ops.Cells.Clear()
    ws_ops.Range("A1").Value = "1日の流れ"
    ops_lines = [
        "18:00頃 Nightly 実行 (Yahoo 売買代金上位300 → coarse/refine → Gap 集計)",
        "朝: SHINSOKU.xlsm を開き Load → Push で候補を反映",
        "候補の Selected / GapRule / DynamicQty を確認・調整",
        "AutoTrade=1, Live=0/1 を確認し Start Auto",
        "DynamicQty = floor(予算 ÷ (価格×(1+slip))) を刻みに丸めて使用",
        "TP/SL はテンプレートに従って指値、Close-Out Time で引け成行",
        "ドライランは Orders/PnL に DEMO として記録",
    ]
    for idx, line in enumerate(ops_lines, start=2):
        ws_ops.Range(f"A{idx}").Value = line
    ws_ops.Columns.AutoFit()


def main():
    if not WB_PATH.exists():
        raise SystemExit(f"Workbook not found: {WB_PATH}")

    excel = ensure_excel()
    wb = None
    try:
        wb = open_workbook(excel)
        clear_helper_sheets(wb)
        ws = ensure_sheet(wb, SHEET_NAME)
        apply_config_defaults(ws)
        write_headers(ws)
        apply_realtime_formulas(ws)
        apply_config_labels(ws)
        ensure_candidates_sheet(wb)
        ensure_orders_sheet(wb)
        seed_ms2_config(wb)
        install_buttons(ws)
        ensure_support_sheets(wb)
        wb.Save()
    finally:
        if wb is not None:
            with contextlib.suppress(Exception):
                wb.Close(SaveChanges=True)
        with contextlib.suppress(Exception):
            excel.Quit()


if __name__ == "__main__":
    main()
