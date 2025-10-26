import argparse
from pathlib import Path

import win32com.client as win32  # type: ignore


ASK_PRICE_LABEL = "\u6700\u826f\u58f2\u6c17\u914d\u5024{level}"
ASK_SIZE_LABEL = "\u6700\u826f\u58f2\u6c17\u914d\u6570\u91cf{level}"
BID_PRICE_LABEL = "\u6700\u826f\u8cb7\u6c17\u914d\u5024{level}"
BID_SIZE_LABEL = "\u6700\u826f\u8cb7\u6c17\u914d\u6570\u91cf{level}"


def set_formula(range_obj, formula: str) -> None:
    range_obj.FormulaLocal = formula


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--board", default=r"excel/BoardLogger.xlsx")
    args = ap.parse_args()

    path = Path(args.board)
    if not path.exists():
        raise SystemExit(f"Board workbook not found: {path}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(path.resolve()))
        names = {name.Name: name for name in wb.Names}

        ticker_ref = names.get("Ticker").RefersToRange if "Ticker" in names else None
        if ticker_ref is None:
            raise RuntimeError("Ticker named range not defined in workbook")

        for i in range(10):
            level = i + 1
            ask_p = names.get(f"ASK_P_{i}")
            ask_q = names.get(f"ASK_Q_{i}")
            bid_p = names.get(f"BID_P_{i}")
            bid_q = names.get(f"BID_Q_{i}")
            if not all([ask_p, ask_q, bid_p, bid_q]):
                raise RuntimeError(f"Missing named range for level {i}")

            ask_p_range = ask_p.RefersToRange
            ask_q_range = ask_q.RefersToRange
            bid_p_range = bid_p.RefersToRange
            bid_q_range = bid_q.RefersToRange

            ask_price_formula = f'=IF(Ticker="","",RssMarket(Ticker,"{ASK_PRICE_LABEL.format(level=level)}"))'
            ask_size_formula = f'=IF(Ticker="","",RssMarket(Ticker,"{ASK_SIZE_LABEL.format(level=level)}"))'
            bid_price_formula = f'=IF(Ticker="","",RssMarket(Ticker,"{BID_PRICE_LABEL.format(level=level)}"))'
            bid_size_formula = f'=IF(Ticker="","",RssMarket(Ticker,"{BID_SIZE_LABEL.format(level=level)}"))'

            set_formula(ask_p_range, ask_price_formula)
            set_formula(ask_q_range, ask_size_formula)
            set_formula(bid_p_range, bid_price_formula)
            set_formula(bid_q_range, bid_size_formula)

        top3 = names.get("TOP3_AMT")
        top10 = names.get("TOP10_AMT")
        if top3:
            top3.RefersToRange.FormulaLocal = (
                "=IF(Ticker=\"\",\"\",SUMPRODUCT($B$5:$B$7,$C$5:$C$7)+SUMPRODUCT($D$5:$D$7,$E$5:$E$7))"
            )
        if top10:
            top10.RefersToRange.FormulaLocal = (
                "=IF(Ticker=\"\",\"\",SUMPRODUCT($B$5:$B$14,$C$5:$C$14)+SUMPRODUCT($D$5:$D$14,$E$5:$E$14))"
            )

        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
