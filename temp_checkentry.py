import win32com.client as win32
from pathlib import Path
path = Path(r"C:\AI\asagake\SHINSOKU.xlsm")
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(str(path))
try:
    ws = wb.Worksheets('NewDashboard')
    excel.Run('AutoTrader.ResetDashboardHeaders')
    excel.Run('AutoTrader.EnsureQueueNowButton')
    excel.Run('AutoTrader.ButtonLoadCandidates')
    excel.Run('AutoTrader.ButtonPushCandidates')
    header_row = 5
    data_start = 6
    def find_column(name):
        used_cols = ws.UsedRange.Columns.Count
        for col in range(1, used_cols + 1):
            value = ws.Cells(header_row, col).Value
            if value is None:
                continue
            if str(value).strip() == name:
                return col
        return None
    ticker_col = find_column('Ticker')
    entry_buy_col = find_column('EntryBuyPx')
    entry_sell_col = find_column('EntrySellPx')
    tp_col = find_column('TP_price')
    sl_col = find_column('SL_price')
    for row in range(data_start, data_start + 10):
        val = ws.Cells(row, ticker_col).Value
        if val:
            print('row', row, 'ticker', val, 'entry_buy', ws.Cells(row, entry_buy_col).Value, 'entry_sell', ws.Cells(row, entry_sell_col).Value, 'tp', ws.Cells(row, tp_col).Value, 'sl', ws.Cells(row, sl_col).Value)
            break
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()
