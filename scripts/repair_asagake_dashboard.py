#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rebuild ASAGAKE NewDashboardV2 using a copy-edit workflow.

Usage:
    python scripts/repair_asagake_dashboard.py --excel C:/AI/asagake/ASAGAKE.xlsm

The script:
  1. Creates a timestamped working copy of the given workbook.
  2. Updates headers, comments, buttons, and formulas on NewDashboardV2.
  3. Saves the working copy, backs up the original workbook, then swaps it in.
"""

from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path


PARAM_HEADERS = [
    "NKY_Code",
    "NKY_Last",
    "NKY_ChgPct",
    "TOPIX_Code",
    "TOPIX_Last",
    "TOPIX_ChgPct",
    "Bias_bp",
    "BiasSlope",
    "GapSlope",
    "GapBanPct",
    "NoTradeMin",
    "TP_per_J",
    "SL_per_J",
    "Trail_per_J",
    "CorrSlope",
    "BudgetPerTicker",
    "LotSize",
    "NKY_TrendDay",
    "NKY_TrendWindow",
    "NKY_AllowedSide",
    "TOPIX_TrendDay",
    "TOPIX_TrendWindow",
    "TOPIX_AllowedSide",
]

PARAM_DEFAULTS = [
    "N225",
    "",
    0.0,
    "TOPX",
    "",
    0.0,
    0.0,
    0.10,
    0.20,
    3.0,
    5,
    0.15,
    0.10,
    0.10,
    0.05,
    1_000_000,
    100,
    "",
    "",
    "BOTH",
    "",
    "",
    "BOTH",
]

PARAM_LABELS = {
    "NKY_Code": ("指標コード(日経平均)", "楽天RSSに渡す指標コード。通常はN225固定です"),
    "NKY_Last": ("日経平均 現在値", "RssIndexMarket(\"現在値\")で取得した値を表示します"),
    "NKY_ChgPct": ("日経平均 前日比", "RssIndexMarket(\"前日比\")の結果を表示します"),
    "TOPIX_Code": ("指標コード(TOPIX)", "楽天RSSに渡すTOPIXコード。通常はTOPXです"),
    "TOPIX_Last": ("TOPIX 現在値", "RssIndexMarket(\"現在値\")で取得したTOPIXの値"),
    "TOPIX_ChgPct": ("TOPIX 前日比", "RssIndexMarket(\"前日比\")の結果"),
    "Bias_bp": ("バイアス閾値(bp)", "J補正をBAN扱いにする絶対値閾値"),
    "BiasSlope": ("Bias補正係数", "日経平均との方向差で加算する係数"),
    "GapSlope": ("Gap補正係数", "ギャップ量に応じた補正係数"),
    "GapBanPct": ("Gap BAN 閾値(%)", "この割合を超えるギャップは自動BAN"),
    "NoTradeMin": ("取引停止分数", "AutoTrader が再開するまでのクールダウン（分）"),
    "TP_per_J": ("TP/J (全体)", "銘柄個別値が空欄の際に参照するTP/J"),
    "SL_per_J": ("SL/J (全体)", "銘柄個別値が空欄の際に参照するSL/J"),
    "Trail_per_J": ("Trail/J (全体)", "銘柄個別値が空欄の際に参照するTrail/J"),
    "CorrSlope": ("相関補正係数", "NKY/TOPIX相関でJ_thを補正する係数"),
    "BudgetPerTicker": ("銘柄別予算(円)", "1銘柄あたりの最大投入額"),
    "LotSize": ("ロットサイズ", "1ポジションあたりの株数"),
    "NKY_TrendDay": ("NKY日足トレンド", "AutoTraderAdvanced算出の方向（BUY/SELL/flat）"),
    "NKY_TrendWindow": ("NKY窓トレンド", "短期窓の方向判定"),
    "NKY_AllowedSide": ("NKY許容サイド", "方向フィルタで許容されるサイド"),
    "TOPIX_TrendDay": ("TOPIX日足トレンド", "TOPIXベースの日足方向"),
    "TOPIX_TrendWindow": ("TOPIX窓トレンド", "TOPIXベースの短期方向"),
    "TOPIX_AllowedSide": ("TOPIX許容サイド", "TOPIXベースで許容される方向"),
}

PARAM_INDEX = {name: idx for idx, name in enumerate(PARAM_HEADERS, start=1)}

JP_MAP = {
    "Ticker": ("証券コード", "候補銘柄のコード。CSVから読み込みます"),
    "Name": ("銘柄名称", "RssMarket(\"銘柄名称\")で自動取得"),
    "J_th_base": ("J_thベース", "週末/ナイトバッチから取り込むベース閾値"),
    "J_th": ("J_th(補正後)", "市場バイアス・ギャップ・相関で補正された閾値"),
    "J": ("J値", "（現在値?VWAP）/ATR_n/100"),
    "Last": ("現在値", "RssMarket(\"現在値\")"),
    "VWAP": ("出来高加重平均", "RssMarket(\"出来高加重平均\")"),
    "OrderQtyPlan": ("発注予定数量", "予算÷現在値をロット単位で丸めた数量"),
    "Selected": ("監視ON/OFF", "1で監視対象"),
    "EntryBuyPx": ("買指値", "VWAPを基準にJ_thで算出"),
    "EntrySellPx": ("売指値", "VWAPを基準にJ_thで算出"),
    "EntrySide": ("サイド", "J値の符号でBUY/SELL"),
    "EntryStatus": ("発注ステータス", "発注処理で書き込み"),
    "TP_price": ("利確指値", "J値×TP_per_J"),
    "SL_price": ("損切指値", "J値×SL_per_J"),
    "StopTrail": ("トレーリング", "トレーリング用セル（必要時にVBAが更新）"),
    "SettleStatus": ("決済状況", "決済ログで使用"),
    "BestBid": ("最良買気配値", "RssMarket(\"最良買気配値\")"),
    "BestAsk": ("最良売気配値", "RssMarket(\"最良売気配値\")"),
    "Gap_bp": ("ギャップ(bp)", "(中値?前日終値)/前日終値×10000"),
    "CorrNKY": ("相関(NKY)", "銘柄と日経平均の相関係数"),
    "PrevClose": ("前日終値", "RssMarket(\"前日終値\")"),
    "ForwardPfEff": ("フォワードPF効率", "週末/ナイト結果"),
    "WinCiLow": ("勝率CI下限", "週末/ナイト結果"),
    "ForwardTrades": ("フォワード取引数", "週末/ナイト結果"),
    "ExpBp": ("フォワード期待bp", "週末/ナイト結果"),
    "ATR_n": ("ATR期間", "バッチ推奨値（無い場合は2）"),
    "TPk": ("TP倍率", "バッチ推奨値"),
    "SLk": ("SL倍率", "バッチ推奨値"),
    "SignalMode": ("シグナルモード", "j-only / j-cross 等"),
    "session": ("取引セッション", "AM0930 等"),
    "plan_tag": ("バッチプラン", "refine のプラン名"),
    "BiasSlope_row": ("Bias係数(銘柄)", "銘柄固有のBiasSlope。空欄なら行2の値を使用"),
    "GapSlope_row": ("Gap係数(銘柄)", "銘柄固有のGapSlope。空欄なら行2の値を使用"),
    "CorrSlope_row": ("Corr係数(銘柄)", "銘柄固有のCorrSlope。空欄なら行2の値を使用"),
    "TP_per_J_row": ("TP/J基準(銘柄)", "銘柄固有のTP/J基準値"),
    "SL_per_J_row": ("SL/J基準(銘柄)", "銘柄固有のSL/J基準値"),
    "Trail_per_J_row": ("Trail/J基準(銘柄)", "銘柄固有のトレーリング幅基準"),
    "TP_per_J_eff": ("TP/J(実効)", "動的調整後のTP/J係数"),
    "SL_per_J_eff": ("SL/J(実効)", "動的調整後のSL/J係数"),
    "Trail_per_J_eff": ("Trail/J(実効)", "動的調整後のトレーリング係数"),
    "BatchKind": ("Batch種別", "nightly / weekend のバッチ区分"),
    "NKY_day_trend": ("NKY日次トレンド", "AutoTraderAdvanced が RSS から算出"),
    "NKY_window_trend": ("NKY窓トレンド", "直近15分回帰＋寄付き比較で判定"),
    "NKY_allowed_side": ("NKY許容サイド", "BUY/SELL/BOTH を表示"),
    "J_ratio": ("J到達率", "|J| / |J_th|"),
    "VolatilityTag": ("ボラタグ", "当日ボラ状況メモ"),
    "trend_driver": ("トレンドドライバ", "NKY / NKY_F / TOPIX のいずれか"),
    "trend_window": ("トレンド窓", "day / window などの判定窓"),
    "trend_bp_th": ("トレンド閾値(bp)", "方向フィルタのbp閾値"),
    "trend_allowed_policy": ("トレンド許容ポリシー", "ALIGNED_ONLY / BOTH など"),
    "driver_day_trend": ("ドライバ日次トレンド", "銘柄ごとの trend_driver に対応した日次方向"),
    "driver_window_trend": ("ドライバ窓トレンド", "銘柄ごとの trend_driver に対応した短期方向"),
    "driver_allowed_side": ("ドライバ許容サイド", "銘柄ごとの trend_driver で許容されたサイド"),
}

BUTTONS = [
    ("btn_live_start", "本番取引開始", 3, 4, "AutoTraderAdvanced.StartLiveV2"),
    ("btn_live_stop", "本番取引停止", 3, 6, "AutoTraderAdvanced.StopLiveV2"),
    ("btn_demo_start", "デモ取引開始", 3, 8, "AutoTraderAdvanced.StartDemoV2"),
    ("btn_demo_stop", "デモ取引停止", 3, 10, "AutoTraderAdvanced.StopDemoV2"),
    ("btn_import", "候補銘柄取込", 3, 12, "AutoTraderAdvanced.ImportCandidatesV2"),
]

ORDER_HEADERS = list(JP_MAP.keys())
COL_INDEX = {name: idx for idx, name in enumerate(ORDER_HEADERS, start=1)}

DATA_START_ROW = 6
DATA_END_ROW = 605
STATUS_RANGE = "A3:B3"


def backup_path(path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_backup_{ts}{path.suffix}")


def col_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def param_ref(name: str) -> str:
    return f"R2C{PARAM_INDEX[name]}"


def rgb(red: int, green: int, blue: int) -> int:
    return red + (green << 8) + (blue << 16)


def set_comment(cell, text: str) -> None:
    try:
        cell.Comment.Delete()
    except Exception:
        pass
    if not text:
        return
    try:
        cell.AddComment(text)
        cell.Comment.Visible = False
    except Exception:
        pass


def create_button(ws, name: str, caption: str, row: int, col: int, macro: str) -> None:
    left = ws.Cells(row, col).Left
    top = ws.Cells(row, col).Top
    width = ws.Range(ws.Cells(row, col), ws.Cells(row, col + 1)).Width - 4
    height = ws.Cells(row, col).RowHeight - 2
    shape = ws.Shapes.AddShape(1, left, top, width, height)
    try:
        shape.Name = name
    except Exception:
        pass
    shape.TextFrame.Characters().Text = caption
    shape.TextFrame.HorizontalAlignment = -4108
    shape.TextFrame.VerticalAlignment = -4108
    shape.OnAction = macro


def rc_relative(current: int, target: int) -> str:
    delta = target - current
    if delta == 0:
        return "RC"
    return f"RC[{delta}]"


def fill_formula(ws, col: int, formula: str) -> None:
    rng = ws.Range(ws.Cells(DATA_START_ROW, col), ws.Cells(DATA_END_ROW, col))
    try:
        rng.FormulaR1C1 = formula
    except Exception:
        rng.FormulaR1C1Local = formula


def build_dashboard(excel_path: Path) -> None:
    import win32com.client  # type: ignore
    from win32com.client import constants  # type: ignore
    xl_expression = getattr(constants, "xlExpression", 2)

    work_copy = excel_path.with_name(f"{excel_path.stem}_work_{datetime.now():%Y%m%d_%H%M%S}{excel_path.suffix}")
    shutil.copy2(excel_path, work_copy)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(work_copy))
        try:
            ws = wb.Worksheets("NewDashboardV2")
        except Exception:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = "NewDashboardV2"

        for idx, title in enumerate(PARAM_HEADERS, start=1):
            jp_label, jp_comment = PARAM_LABELS.get(title, (title, ""))
            ws.Cells(1, idx).Value = jp_label
            if jp_comment:
                set_comment(ws.Cells(1, idx), jp_comment)
            force_default = title in {"NKY_Code", "TOPIX_Code"}
            if force_default:
                ws.Cells(2, idx).Value = PARAM_DEFAULTS[idx - 1]
            elif ws.Cells(2, idx).Value in ("", None):
                ws.Cells(2, idx).Value = PARAM_DEFAULTS[idx - 1]

        # Ensure NKY related formulas persist (RSS index depends on code in A2)
        try:
            ws.Cells(2, 2).FormulaR1C1Local = '=IF(RC[-1]="", "", IFERROR(RssIndexMarket(RC[-1],"現在値"),""))'
            ws.Cells(2, 3).FormulaR1C1Local = '=IF(RC[-2]="", "", IFERROR(RssIndexMarket(RC[-2],"前日比"),""))'
        except Exception:
            pass
        try:
            topix_code_idx = PARAM_HEADERS.index("TOPIX_Code") + 1
            topix_last_idx = PARAM_HEADERS.index("TOPIX_Last") + 1
            topix_chg_idx = PARAM_HEADERS.index("TOPIX_ChgPct") + 1
            ws.Cells(2, topix_last_idx).FormulaR1C1Local = '=IF(RC[-1]="", "", IFERROR(RssIndexMarket(RC[-1],"現在値"),""))'
            ws.Cells(2, topix_chg_idx).FormulaR1C1Local = '=IF(RC[-2]="", "", IFERROR(RssIndexMarket(RC[-2],"前日比"),""))'
        except Exception:
            pass

        status_rng = ws.Range(STATUS_RANGE)
        status_rng.Merge()
        status_rng.HorizontalAlignment = -4108
        status_rng.VerticalAlignment = -4108
        status_rng.Font.Bold = True
        status_rng.Font.Size = 16
        status_rng.Interior.Color = 15132390
        status_rng.Value = "IDLE"
        try:
            status_rng.Name = "RunStatusV2"
        except Exception:
            pass

        for shape in list(ws.Shapes):
            if shape.Name.lower().startswith("btn_"):
                shape.Delete()
        for name, caption, row, col, macro in BUTTONS:
            create_button(ws, name, caption, row, col, macro)

        last_col = col_letter(len(ORDER_HEADERS))
        for idx, header in enumerate(ORDER_HEADERS, start=1):
            jp, comment = JP_MAP.get(header, (header, ""))
            ws.Cells(4, idx).Value = jp
            ws.Cells(5, idx).Value = header
            set_comment(ws.Cells(4, idx), comment)

        ws.Range(f"A4:{last_col}5").Interior.Color = 15790320
        ws.Range(f"A5:{last_col}5").Font.Bold = True

        ws.Activate()
        excel.ActiveWindow.SplitRow = 5
        excel.ActiveWindow.FreezePanes = True

        ticker_col = COL_INDEX["Ticker"]

        fill_formula(
            ws,
            COL_INDEX["Name"],
            f'=IF({rc_relative(COL_INDEX["Name"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["Name"], ticker_col)},".T",""),"銘柄名称"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["Last"],
            f'=IF({rc_relative(COL_INDEX["Last"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["Last"], ticker_col)},".T",""),"現在値"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["VWAP"],
            f'=IF({rc_relative(COL_INDEX["VWAP"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["VWAP"], ticker_col)},".T",""),"出来高加重平均"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["PrevClose"],
            f'=IF({rc_relative(COL_INDEX["PrevClose"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["PrevClose"], ticker_col)},".T",""),"前日終値"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["BestBid"],
            f'=IF({rc_relative(COL_INDEX["BestBid"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["BestBid"], ticker_col)},".T",""),"最良買気配値"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["BestAsk"],
            f'=IF({rc_relative(COL_INDEX["BestAsk"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["BestAsk"], ticker_col)},".T",""),"最良売気配値"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["Gap_bp"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["BestBid"])}="",'
                f'{rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["BestAsk"])}="",'
                f'{rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["PrevClose"])}="",'
                f'{rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["PrevClose"])}=0),"",'
                f'(({rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["BestBid"])}+{rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["BestAsk"])})/2-'
                f'{rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["PrevClose"])})/'
                f'{rc_relative(COL_INDEX["Gap_bp"], COL_INDEX["PrevClose"])}*10000)'
            ),
        )

        fill_formula(
            ws,
            COL_INDEX["J_th"],
            (
                f'=IF(OR(ABS(N({rc_relative(COL_INDEX["J_th"], COL_INDEX["Gap_bp"])}))/100>{param_ref("Bias_bp")},'
                f'AND({rc_relative(COL_INDEX["J_th"], COL_INDEX["trend_allowed_policy"])}<>"",'
                f'UPPER({rc_relative(COL_INDEX["J_th"], COL_INDEX["trend_allowed_policy"])})="ALIGNED_ONLY",'
                f'{rc_relative(COL_INDEX["J_th"], COL_INDEX["EntrySide"])}<>"",'
                f'{rc_relative(COL_INDEX["J_th"], COL_INDEX["driver_allowed_side"])}<>"",'
                f'UPPER({rc_relative(COL_INDEX["J_th"], COL_INDEX["driver_allowed_side"])})<>"BOTH",'
                f'UPPER({rc_relative(COL_INDEX["J_th"], COL_INDEX["driver_allowed_side"])})<>UPPER({rc_relative(COL_INDEX["J_th"], COL_INDEX["EntrySide"])})'
                f')), "BAN",'
                f'{rc_relative(COL_INDEX["J_th"], COL_INDEX["J_th_base"])}+'
                f'IF({rc_relative(COL_INDEX["J_th"], COL_INDEX["BiasSlope_row"])}="",{param_ref("BiasSlope")},{rc_relative(COL_INDEX["J_th"], COL_INDEX["BiasSlope_row"])})*{param_ref("NKY_ChgPct")}/100+'
                f'IF({rc_relative(COL_INDEX["J_th"], COL_INDEX["GapSlope_row"])}="",{param_ref("GapSlope")},{rc_relative(COL_INDEX["J_th"], COL_INDEX["GapSlope_row"])})*ABS(N({rc_relative(COL_INDEX["J_th"], COL_INDEX["Gap_bp"])}))/100+'
                f'IF({rc_relative(COL_INDEX["J_th"], COL_INDEX["CorrSlope_row"])}="",{param_ref("CorrSlope")},{rc_relative(COL_INDEX["J_th"], COL_INDEX["CorrSlope_row"])})*N({rc_relative(COL_INDEX["J_th"], COL_INDEX["CorrNKY"])})*{param_ref("NKY_ChgPct")}/100)'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["J"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["J"], COL_INDEX["Last"])}="",'
                f'{rc_relative(COL_INDEX["J"], COL_INDEX["VWAP"])}=""),"",'
                f'(({rc_relative(COL_INDEX["J"], COL_INDEX["Last"])}-'
                f'{rc_relative(COL_INDEX["J"], COL_INDEX["VWAP"])})/'
                f'IF(N({rc_relative(COL_INDEX["J"], COL_INDEX["ATR_n"])})=0,2,N({rc_relative(COL_INDEX["J"], COL_INDEX["ATR_n"])})))/100)'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["Last"],
            (
                f'=IF({rc_relative(COL_INDEX["Last"], COL_INDEX["Ticker"])}="", "",'
                f'IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["Last"], COL_INDEX["Ticker"])},".T",""),"現在値"),""))'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["OrderQtyPlan"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["OrderQtyPlan"], COL_INDEX["Last"])}="",'
                f'{rc_relative(COL_INDEX["OrderQtyPlan"], COL_INDEX["Last"])}=0), "",'
                f'MAX(0,INT(R2C13/{rc_relative(COL_INDEX["OrderQtyPlan"], COL_INDEX["Last"])}/R2C14)*R2C14))'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["EntryBuyPx"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["J_th"])}="BAN",'
                f'{rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["J_th"])}="",'
                f'{rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["Last"])}=""),"",'
                f'IF({rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["VWAP"])}="",'
                f'{rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["PrevClose"])},'
                f'{rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["VWAP"])})'
                f'-0.001*ABS(N({rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["J_th"])}))*'
                f'IF({rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["VWAP"])}="",'
                f'{rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["PrevClose"])},'
                f'{rc_relative(COL_INDEX["EntryBuyPx"], COL_INDEX["VWAP"])})'
                ')'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["EntrySellPx"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["J_th"])}="BAN",'
                f'{rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["J_th"])}="",'
                f'{rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["Last"])}=""),"",'
                f'IF({rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["VWAP"])}="",'
                f'{rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["PrevClose"])},'
                f'{rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["VWAP"])})'
                f'+0.001*ABS(N({rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["J_th"])}))*'
                f'IF({rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["VWAP"])}="",'
                f'{rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["PrevClose"])},'
                f'{rc_relative(COL_INDEX["EntrySellPx"], COL_INDEX["VWAP"])})'
                ')'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["EntrySide"],
            (
                f'=IF({rc_relative(COL_INDEX["EntrySide"], COL_INDEX["J"])}<0,"BUY",'
                f'IF({rc_relative(COL_INDEX["EntrySide"], COL_INDEX["J"])}>0,"SELL",""))'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["TP_price"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["TP_price"], COL_INDEX["J_th"])}="BAN",'
                f'{rc_relative(COL_INDEX["TP_price"], COL_INDEX["EntrySide"])}=""),"",'
                f'IF({rc_relative(COL_INDEX["TP_price"], COL_INDEX["EntrySide"])}="BUY",'
                f'{rc_relative(COL_INDEX["TP_price"], COL_INDEX["Last"])}*(1+N(IF({rc_relative(COL_INDEX["TP_price"], COL_INDEX["TP_per_J_eff"])}="",R2C9,'
                f'{rc_relative(COL_INDEX["TP_price"], COL_INDEX["TP_per_J_eff"])}))*ABS({rc_relative(COL_INDEX["TP_price"], COL_INDEX["J"])})/100),'
                f'IF({rc_relative(COL_INDEX["TP_price"], COL_INDEX["EntrySide"])}="SELL",'
                f'{rc_relative(COL_INDEX["TP_price"], COL_INDEX["Last"])}*(1-N(IF({rc_relative(COL_INDEX["TP_price"], COL_INDEX["TP_per_J_eff"])}="",R2C9,'
                f'{rc_relative(COL_INDEX["TP_price"], COL_INDEX["TP_per_J_eff"])}))*ABS({rc_relative(COL_INDEX["TP_price"], COL_INDEX["J"])})/100),"")))'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["SL_price"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["SL_price"], COL_INDEX["J_th"])}="BAN",'
                f'{rc_relative(COL_INDEX["SL_price"], COL_INDEX["EntrySide"])}=""),"",'
                f'IF({rc_relative(COL_INDEX["SL_price"], COL_INDEX["EntrySide"])}="BUY",'
                f'{rc_relative(COL_INDEX["SL_price"], COL_INDEX["Last"])}*(1-N(IF({rc_relative(COL_INDEX["SL_price"], COL_INDEX["SL_per_J_eff"])}="",R2C10,'
                f'{rc_relative(COL_INDEX["SL_price"], COL_INDEX["SL_per_J_eff"])}))*ABS({rc_relative(COL_INDEX["SL_price"], COL_INDEX["J"])})/100),'
                f'IF({rc_relative(COL_INDEX["SL_price"], COL_INDEX["EntrySide"])}="SELL",'
                f'{rc_relative(COL_INDEX["SL_price"], COL_INDEX["Last"])}*(1+N(IF({rc_relative(COL_INDEX["SL_price"], COL_INDEX["SL_per_J_eff"])}="",R2C10,'
                f'{rc_relative(COL_INDEX["SL_price"], COL_INDEX["SL_per_J_eff"])}))*ABS({rc_relative(COL_INDEX["SL_price"], COL_INDEX["J"])})/100),"")))'
            ),
        )
        fill_formula(
            ws,
            COL_INDEX["StopTrail"],
            (
                f'=IF(OR({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["EntrySide"])}="",'
                f'{rc_relative(COL_INDEX["StopTrail"], COL_INDEX["J_th"])}="BAN"),"",'
                f'IF({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["EntrySide"])}="BUY",'
                f'{rc_relative(COL_INDEX["StopTrail"], COL_INDEX["Last"])}*(1-N(IF({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["Trail_per_J_eff"])}="",R2C11,'
                f'{rc_relative(COL_INDEX["StopTrail"], COL_INDEX["Trail_per_J_eff"])}))*ABS({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["J"])})/100),'
                f'IF({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["EntrySide"])}="SELL",'
                f'{rc_relative(COL_INDEX["StopTrail"], COL_INDEX["Last"])}*(1+N(IF({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["Trail_per_J_eff"])}="",R2C11,'
                f'{rc_relative(COL_INDEX["StopTrail"], COL_INDEX["Trail_per_J_eff"])}))*ABS({rc_relative(COL_INDEX["StopTrail"], COL_INDEX["J"])})/100),"")))'
            ),
        )

        selected_col = COL_INDEX["Selected"]
        atr_col = COL_INDEX["ATR_n"]
        for row in range(DATA_START_ROW, DATA_END_ROW + 1):
            if ws.Cells(row, selected_col).Value in ("", None):
                ws.Cells(row, selected_col).Value = 1
            if ws.Cells(row, atr_col).Value in ("", None):
                ws.Cells(row, atr_col).Value = 2

        ratio_letter = col_letter(COL_INDEX["J_ratio"])
        try:
            ratio_range = ws.Range(
                f"{ratio_letter}{DATA_START_ROW}:{ratio_letter}{DATA_END_ROW}"
            )
            ratio_range.FormatConditions.Delete()
            cond_high = ratio_range.FormatConditions.Add(
                Type=xl_expression,
                Formula1=f"=ABS(${ratio_letter}{DATA_START_ROW})>=1"
            )
            cond_high.Interior.Color = rgb(146, 208, 80)
            cond_high.StopIfTrue = False
            cond_mid = ratio_range.FormatConditions.Add(
                Type=xl_expression,
                Formula1=f"=AND(ABS(${ratio_letter}{DATA_START_ROW})>=0.8,ABS(${ratio_letter}{DATA_START_ROW})<1)"
            )
            cond_mid.Interior.Color = rgb(198, 239, 206)
            cond_mid.StopIfTrue = False
        except Exception:
            pass

        allowed_letter = col_letter(COL_INDEX["driver_allowed_side"])
        entry_letter = col_letter(COL_INDEX["EntrySide"])
        policy_letter = col_letter(COL_INDEX["trend_allowed_policy"])
        try:
            row_range = ws.Range(f"A{DATA_START_ROW}:{last_col}{DATA_END_ROW}")
            row_range.FormatConditions.Delete()
            gray_formula = (
                f'=AND(${policy_letter}{DATA_START_ROW}<>"",'
                f'UPPER(${policy_letter}{DATA_START_ROW})="ALIGNED_ONLY",'
                f'${entry_letter}{DATA_START_ROW}<>"",'
                f'${allowed_letter}{DATA_START_ROW}<>"",'
                f'UPPER(${allowed_letter}{DATA_START_ROW})<>"BOTH",'
                f'UPPER(${allowed_letter}{DATA_START_ROW})<>UPPER(${entry_letter}{DATA_START_ROW}))'
            )
            cond_gray = row_range.FormatConditions.Add(
                Type=xl_expression,
                Formula1=gray_formula,
            )
            cond_gray.Interior.Color = rgb(242, 242, 242)
            cond_gray.Font.Color = rgb(109, 109, 109)
            cond_gray.StopIfTrue = False
        except Exception:
            pass

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()

    backup = backup_path(excel_path)
    shutil.copy2(excel_path, backup)
    shutil.copy2(work_copy, excel_path)
    work_copy.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default=r"C:/AI/asagake/ASAGAKE.xlsm")
    args = parser.parse_args()
    build_dashboard(Path(args.excel).resolve())


if __name__ == "__main__":
    main()
