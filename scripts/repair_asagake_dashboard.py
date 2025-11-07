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
]

PARAM_DEFAULTS = [
    "N225",
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
]

PARAM_INDEX = {name: idx for idx, name in enumerate(PARAM_HEADERS, start=1)}

JP_MAP = {
    "Ticker": ("險ｼ蛻ｸ繧ｳ繝ｼ繝・, "蛟呵｣憺釜譟・・繧ｳ繝ｼ繝峨・SV縺九ｉ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺・),
    "Name": ("驫俶氛蜷咲ｧｰ", "RssMarket(\"驫俶氛蜷咲ｧｰ\")縺ｧ閾ｪ蜍募叙蠕・),
    "J_th_base": ("J_th繝吶・繧ｹ", "騾ｱ譛ｫ/繝翫う繝医ヰ繝・メ縺九ｉ蜿悶ｊ霎ｼ繧繝吶・繧ｹ髢ｾ蛟､"),
    "J_th": ("J_th(陬懈ｭ｣蠕・", "蟶ょｴ繝舌う繧｢繧ｹ繝ｻ繧ｮ繝｣繝・・繝ｻ逶ｸ髢｢縺ｧ陬懈ｭ｣縺輔ｌ縺滄明蛟､"),
    "J": ("J蛟､", "・育樟蝨ｨ蛟､竏歎WAP・・ATR_n/100"),
    "Last": ("迴ｾ蝨ｨ蛟､", "RssMarket(\"迴ｾ蝨ｨ蛟､\")"),
    "VWAP": ("蜃ｺ譚･鬮伜刈驥榊ｹｳ蝮・, "RssMarket(\"蜃ｺ譚･鬮伜刈驥榊ｹｳ蝮Ⅸ")"),
    "OrderQtyPlan": ("逋ｺ豕ｨ莠亥ｮ壽焚驥・, "莠育ｮ療ｷ迴ｾ蝨ｨ蛟､繧偵Ο繝・ヨ蜊倅ｽ阪〒荳ｸ繧√◆謨ｰ驥・),
    "Selected": ("逶｣隕飽N/OFF", "1縺ｧ逶｣隕門ｯｾ雎｡"),
    "EntryBuyPx": ("雋ｷ謖・､", "VWAP繧貞渕貅悶↓J_th縺ｧ邂怜・"),
    "EntrySellPx": ("螢ｲ謖・､", "VWAP繧貞渕貅悶↓J_th縺ｧ邂怜・"),
    "EntrySide": ("繧ｵ繧､繝・, "J蛟､縺ｮ隨ｦ蜿ｷ縺ｧBUY/SELL"),
    "EntryStatus": ("逋ｺ豕ｨ繧ｹ繝・・繧ｿ繧ｹ", "逋ｺ豕ｨ蜃ｦ逅・〒譖ｸ縺崎ｾｼ縺ｿ"),
    "TP_price": ("蛻ｩ遒ｺ謖・､", "J蛟､ﾃ裕P_per_J"),
    "SL_price": ("謳榊・謖・､", "J蛟､ﾃ祐L_per_J"),
    "StopTrail": ("繝医Ξ繝ｼ繝ｪ繝ｳ繧ｰ", "繝医Ξ繝ｼ繝ｪ繝ｳ繧ｰ逕ｨ繧ｻ繝ｫ・亥ｿ・ｦ∵凾縺ｫVBA縺梧峩譁ｰ・・),
    "SettleStatus": ("豎ｺ貂育憾豕・, "豎ｺ貂医Ο繧ｰ縺ｧ菴ｿ逕ｨ"),
    "BestBid": ("譛濶ｯ雋ｷ豌鈴・蛟､", "RssMarket(\"譛濶ｯ雋ｷ豌鈴・蛟､\")"),
    "BestAsk": ("譛濶ｯ螢ｲ豌鈴・蛟､", "RssMarket(\"譛濶ｯ螢ｲ豌鈴・蛟､\")"),
    "Gap_bp": ("繧ｮ繝｣繝・・(bp)", "(荳ｭ蛟､竏貞燕譌･邨ょ､)/蜑肴律邨ょ､ﾃ・0000"),
    "CorrNKY": ("逶ｸ髢｢(NKY)", "驫俶氛縺ｨ譌･邨悟ｹｳ蝮・・逶ｸ髢｢菫よ焚"),
    "PrevClose": ("蜑肴律邨ょ､", "RssMarket(\"蜑肴律邨ょ､\")"),
    "ForwardPfEff": ("繝輔か繝ｯ繝ｼ繝臼F蜉ｹ邇・, "騾ｱ譛ｫ/繝翫う繝育ｵ先棡"),
    "WinCiLow": ("蜍晉紫CI荳矩剞", "騾ｱ譛ｫ/繝翫う繝育ｵ先棡"),
    "ForwardTrades": ("繝輔か繝ｯ繝ｼ繝牙叙蠑墓焚", "騾ｱ譛ｫ/繝翫う繝育ｵ先棡"),
    "ExpBp": ("繝輔か繝ｯ繝ｼ繝画悄蠕・p", "騾ｱ譛ｫ/繝翫う繝育ｵ先棡"),
    "ATR_n": ("ATR譛滄俣", "繝舌ャ繝∵耳螂ｨ蛟､・育┌縺・ｴ蜷医・2・・),
    "TPk": ("TP蛟咲紫", "繝舌ャ繝∵耳螂ｨ蛟､"),
    "SLk": ("SL蛟咲紫", "繝舌ャ繝∵耳螂ｨ蛟､"),
    "SignalMode": ("繧ｷ繧ｰ繝翫Ν繝｢繝ｼ繝・, "j-only / j-cross 遲・),
    "session": ("蜿門ｼ輔そ繝・す繝ｧ繝ｳ", "AM0930 遲・),
    "plan_tag": ("繝舌ャ繝√・繝ｩ繝ｳ", "refine 縺ｮ繝励Λ繝ｳ蜷・),
    "BiasSlope_row": ("Bias菫よ焚(驫俶氛)", "驫俶氛蝗ｺ譛峨・BiasSlope縲らｩｺ谺・↑繧芽｡・縺ｮ蛟､繧剃ｽｿ逕ｨ"),
    "GapSlope_row": ("Gap菫よ焚(驫俶氛)", "驫俶氛蝗ｺ譛峨・GapSlope縲らｩｺ谺・↑繧芽｡・縺ｮ蛟､繧剃ｽｿ逕ｨ"),
    "CorrSlope_row": ("Corr菫よ焚(驫俶氛)", "驫俶氛蝗ｺ譛峨・CorrSlope縲らｩｺ谺・↑繧芽｡・縺ｮ蛟､繧剃ｽｿ逕ｨ"),
    "TP_per_J_row": ("TP/J蝓ｺ貅・驫俶氛)", "驫俶氛蝗ｺ譛峨・TP/J蝓ｺ貅門､"),
    "SL_per_J_row": ("SL/J蝓ｺ貅・驫俶氛)", "驫俶氛蝗ｺ譛峨・SL/J蝓ｺ貅門､"),
    "Trail_per_J_row": ("Trail/J蝓ｺ貅・驫俶氛)", "驫俶氛蝗ｺ譛峨・繝医Ξ繝ｼ繝ｪ繝ｳ繧ｰ蟷・渕貅・),
    "TP_per_J_eff": ("TP/J(螳溷柑)", "蜍慕噪隱ｿ謨ｴ蠕後・TP/J菫よ焚"),
    "SL_per_J_eff": ("SL/J(螳溷柑)", "蜍慕噪隱ｿ謨ｴ蠕後・SL/J菫よ焚"),
    "Trail_per_J_eff": ("Trail/J(螳溷柑)", "蜍慕噪隱ｿ謨ｴ蠕後・繝医Ξ繝ｼ繝ｪ繝ｳ繧ｰ菫よ焚"),
    "BatchKind": ("Batch種別", "nightly / weekend などの区別"),
    "NKY_day_trend": ("NKY日次トレンド", "AutoTraderAdvanced が更新"),
    "NKY_window_trend": ("NKY窓トレンド", "直近15本の方向"),
    "NKY_allowed_side": ("NKY許容サイド", "BUY / SELL / BOTH"),
    "J_ratio": ("J到達率", "|J| / |J_th|"),
    "VolatilityTag": ("繝懊Λ繧ｿ繧ｰ", "蠖捺律繝懊Λ迥ｶ豕√Γ繝｢"),
}

BUTTONS = [
    ("btn_live_start", "譛ｬ逡ｪ蜿門ｼ暮幕蟋・, 3, 4, "AutoTraderAdvanced.StartLiveV2"),
    ("btn_live_stop", "譛ｬ逡ｪ蜿門ｼ募●豁｢", 3, 6, "AutoTraderAdvanced.StopLiveV2"),
    ("btn_demo_start", "繝・Δ蜿門ｼ暮幕蟋・, 3, 8, "AutoTraderAdvanced.StartDemoV2"),
    ("btn_demo_stop", "繝・Δ蜿門ｼ募●豁｢", 3, 10, "AutoTraderAdvanced.StopDemoV2"),
    ("btn_import", "蛟呵｣憺釜譟・叙霎ｼ", 3, 12, "AutoTraderAdvanced.ImportCandidatesV2"),
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
            ws.Cells(1, idx).Value = title
            if ws.Cells(2, idx).Value in ("", None):
                ws.Cells(2, idx).Value = PARAM_DEFAULTS[idx - 1]

        # Ensure NKY related formulas persist (RSS index depends on code in A2)
        try:
            ws.Cells(2, 2).FormulaR1C1Local = '=IF(RC[-1]="", "", IFERROR(RssIndexMarket(RC[-1],"迴ｾ蝨ｨ蛟､"),""))'
            ws.Cells(2, 3).FormulaR1C1Local = '=IF(RC[-2]="", "", IFERROR(RssIndexMarket(RC[-2],"鬨ｰ關ｽ邇・),""))'
            ws.Cells(2, 4).FormulaR1C1 = "=(RC[-1])*100"
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
            f'=IF({rc_relative(COL_INDEX["Name"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["Name"], ticker_col)},".T",""),"驫俶氛蜷咲ｧｰ"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["Last"],
            f'=IF({rc_relative(COL_INDEX["Last"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["Last"], ticker_col)},".T",""),"迴ｾ蝨ｨ蛟､"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["VWAP"],
            f'=IF({rc_relative(COL_INDEX["VWAP"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["VWAP"], ticker_col)},".T",""),"蜃ｺ譚･鬮伜刈驥榊ｹｳ蝮・),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["PrevClose"],
            f'=IF({rc_relative(COL_INDEX["PrevClose"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["PrevClose"], ticker_col)},".T",""),"蜑肴律邨ょ､"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["BestBid"],
            f'=IF({rc_relative(COL_INDEX["BestBid"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["BestBid"], ticker_col)},".T",""),"譛濶ｯ雋ｷ豌鈴・蛟､"),""))',
        )
        fill_formula(
            ws,
            COL_INDEX["BestAsk"],
            f'=IF({rc_relative(COL_INDEX["BestAsk"], ticker_col)}="","",IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["BestAsk"], ticker_col)},".T",""),"譛濶ｯ螢ｲ豌鈴・蛟､"),""))',
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
                f'=IF(ABS(N({rc_relative(COL_INDEX["J_th"], COL_INDEX["Gap_bp"])}))/100>R2C7,"BAN",'
                f'{rc_relative(COL_INDEX["J_th"], COL_INDEX["J_th_base"])}+'
                f'IF({rc_relative(COL_INDEX["J_th"], COL_INDEX["BiasSlope_row"])}="",R2C5,{rc_relative(COL_INDEX["J_th"], COL_INDEX["BiasSlope_row"])})*R2C4/100+'
                f'IF({rc_relative(COL_INDEX["J_th"], COL_INDEX["GapSlope_row"])}="",R2C6,{rc_relative(COL_INDEX["J_th"], COL_INDEX["GapSlope_row"])})*ABS(N({rc_relative(COL_INDEX["J_th"], COL_INDEX["Gap_bp"])}))/100+'
                f'IF({rc_relative(COL_INDEX["J_th"], COL_INDEX["CorrSlope_row"])}="",R2C12,{rc_relative(COL_INDEX["J_th"], COL_INDEX["CorrSlope_row"])})*N({rc_relative(COL_INDEX["J_th"], COL_INDEX["CorrNKY"])})*R2C4/100)'
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
                f'IFERROR(RssMarket(SUBSTITUTE({rc_relative(COL_INDEX["Last"], COL_INDEX["Ticker"])},".T",""),"迴ｾ蝨ｨ蛟､"),""))'
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
