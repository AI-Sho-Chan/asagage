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

HEADER_OFFSET_COL = 24
HEADERS = [
    "Selected", "SignalMode", "Session", "ATR_n", "TPk", "SLk", "J_th",
    "ForwardPF", "ForwardTrades", "WinCI_L", "WinCI_H", "ExpBootMean",
    "ExpBootLow", "ExpBootHigh"
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

    # Additional defaults for live/bridge/close-out/qty/tif
    ws.Range("A6").Value = "Selected Default (0/1)"
    if ws.Range("B6").Value in (None, ""):
        ws.Range("B6").Value = 1
    ws.Range("A7").Value = "Reentry Allowed (0/1)"
    if ws.Range("B7").Value in (None, ""):
        ws.Range("B7").Value = 0
    ws.Range("A8").Value = "Hard Stop (0/1)"
    if ws.Range("B8").Value in (None, ""):
        ws.Range("B8").Value = 0
    ws.Range("A9").Value = "Live Orders (0/1)"
    if ws.Range("B9").Value in (None, ""):
        ws.Range("B9").Value = 0
    ws.Range("A10").Value = "Order Macro Name"
    if ws.Range("B10").Value in (None, ""):
        ws.Range("B10").Value = "MS2Bridge.Place"
    ws.Range("A11").Value = "Close-Out Time (HH:MM:SS)"
    if ws.Range("B11").Value in (None, ""):
        ws.Range("B11").Value = "14:59:30"
    ws.Range("A12").Value = "Order Quantity"
    if ws.Range("B12").Value in (None, ""):
        ws.Range("B12").Value = 100
    ws.Range("A13").Value = "TIF/Type"
    if ws.Range("B13").Value in (None, ""):
        ws.Range("B13").Value = "MKT"

    # Create target sheets
    for sheet_name in ("Candidates", "Orders", "MS2_Config"):
        try:
            wb.Worksheets(sheet_name)
        except Exception:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = sheet_name
    ws_orders = wb.Worksheets("Orders")
    if ws_orders.Cells(1, 1).Value in (None, ""):
        ws_orders.Range("A1:E1").Value = ("Time", "Ticker", "Side", "Price", "Note")

    # Seed MS2_Config with keys and sample placeholders
    ws_cfg = wb.Worksheets("MS2_Config")
    if ws_cfg.Cells(1,1).Value in (None,""):
        ws_cfg.Range("A1:C1").Value = ("Key","Value","Notes")
        rows = [
            ("Account", "", "口座識別（必要なら）。"),
            ("Market", "TSE", "市場コード/名称（必要なら）。"),
            ("EntryTemplate", "RssStockOrder({Account},\"{Ticker}\",\"{Side}\",{Qty},{Price},\"{TIF}\")", "実関数・引数に合わせて編集"),
            ("TPTemplate", "RssStockOrder({Account},\"{Ticker}\",\"{Side}\",{Qty},{Price},\"{TIF}\")", "利確用テンプレート"),
            ("SLTemplate", "RssStockOrder({Account},\"{Ticker}\",\"{Side}\",{Qty},{Price},\"{TIF}\")", "損切用テンプレート"),
            ("MOCTemplate", "RssStockOrder({Account},\"{Ticker}\",\"{Side}\",{Qty},,\"MOC\")", "引け成行テンプレート"),
        ]
        r0 = 2
        for i,(k,v,n) in enumerate(rows):
            ws_cfg.Cells(r0+i,1).Value = k
            ws_cfg.Cells(r0+i,2).Value = v
            ws_cfg.Cells(r0+i,3).Value = n
        ws_cfg.Columns.AutoFit()

    # Remove existing buttons by name
    for btn_name, _, _, _ in BUTTONS:
        for shp in list(ws.Shapes):
            if shp.Name == btn_name:
                shp.Delete()

    # Add buttons stacked below header row (row 1 area)
    left_start = ws.Cells(1, 10).Left
    top_start = ws.Cells(1, 1).Top
    width = 120
    height = 24
    spacing = 6
    for idx, (btn_name, caption, macro, order) in enumerate(BUTTONS):
        btn = ws.Buttons().Add(left_start, top_start + (height + spacing) * idx, width, height)
        btn.Name = btn_name
        btn.OnAction = macro
        btn.Text = caption

    wb.Save()
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()

