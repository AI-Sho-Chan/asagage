import argparse
from pathlib import Path


JP_LABELS = [
    "\u8a3c\u5238\u30b3\u30fc\u30c9","\u9298\u67c4\u540d","J\u95be\u5024","J\u5024","\u524d\u65e5\u7d42\u5024","VWAP","\u767a\u6ce8\u4e88\u5b9a\u6570\u91cf","\u76e3\u8996ON/OFF",
    "\u8cb7\u6307\u5024","\u58f2\u6307\u5024","\u30b5\u30a4\u30c9","\u767a\u6ce8\u72b6\u6cc1","\u5229\u78ba\u6307\u5024","\u640d\u5207\u6307\u5024",
    "\u30c8\u30ec\u30fc\u30ea\u30f3\u30b0","\u6c7a\u6e08\u72b6\u6cc1","\u6700\u826f\u8cb7\u6c17\u914d\u5024","\u6700\u826f\u58f2\u6c17\u914d\u5024","\u30ae\u30e3\u30c3\u30d7(bp)","\u76f8\u95a2(NKY)"
]

EN_HEADERS = [
    "Ticker","Name","J_th","J","PrevClose","VWAP","OrderQtyPlan","Selected",
    "EntryBuyPx","EntrySellPx","EntrySide","EntryStatus","TP_price","SL_price",
    "StopTrail","SettleStatus","BestBid","BestAsk","Gap_bp","CorrNKY"
]

JP_EXPLAIN = {
    "\u8a3c\u5238\u30b3\u30fc\u30c9": "驫俶氛繧ｳ繝ｼ繝会ｼ井ｾ・ 7203.T・・,
    "\u9298\u67c4\u540d": "驫俶氛蜷搾ｼ・SS縺九ｉ閾ｪ蜍募叙蠕暦ｼ・,
    "J\u95be\u5024": "繧ｷ繧ｰ繝翫Ν逋ｺ轣ｫ縺ｮ蝓ｺ貅悶→縺ｪ繧徽縺ｮ髢ｾ蛟､・郁｣懈ｭ｣蠕後・閾ｪ蜍戊ｨ育ｮ暦ｼ・,
    "J\u5024": "迴ｾ蝨ｨ縺ｮJ蛟､・・繧剃ｸｭ蠢・↓ﾂｱ譁ｹ蜷代↓荵夜屬繧堤､ｺ縺呻ｼ・,
    "\u524d\u65e5\u7d42\u5024": "蜑肴律邨ょ､・域ｯ碑ｼ・・繧ｮ繝｣繝・・險育ｮ励↓菴ｿ逕ｨ・・,
    "VWAP": "蠖捺律縺ｮVWAP・医お繝ｳ繝医Μ蝓ｺ貅紋ｾ｡譬ｼ・・,
    "\u767a\u6ce8\u4e88\u5b9a\u6570\u91cf": "逋ｺ豕ｨ莠亥ｮ壽焚驥擾ｼ域怙驕ｩ蛹悶い繝ｫ繧ｴ繝ｪ繧ｺ繝縺ｧ邂怜・莠亥ｮ夲ｼ・,
    "\u76e3\u8996ON/OFF": "1=逶｣隕門ｯｾ雎｡・亥呵｣懶ｼ峨・=辟｡蜉ｹ",
    "\u8cb7\u6307\u5024": "雋ｷ縺・・莠句燕謖・､・・縺ｫ蠢懊§縺ｦ閾ｪ蜍戊ｨ育ｮ暦ｼ・,
    "\u58f2\u6307\u5024": "螢ｲ繧翫・莠句燕謖・､・・縺ｫ蠢懊§縺ｦ閾ｪ蜍戊ｨ育ｮ暦ｼ・,
    "\u30b5\u30a4\u30c9": "BUY/SELL縺ｮ繧ｵ繧､繝画耳螳夲ｼ・縺ｮ隨ｦ蜿ｷ縺ｧ豎ｺ螳夲ｼ・,
    "\u767a\u6ce8\u72b6\u6cc1": "preplace/PENDING/邏・ｮ・遲峨・迥ｶ諷・,
    "\u5229\u78ba\u6307\u5024": "蛻ｩ遒ｺ謖・､・・驕主臆蛻・↓蠢懊§縺ｦ蜍慕噪・・,
    "\u640d\u5207\u6307\u5024": "謳榊・謖・､・・驕主臆蛻・↓蠢懊§縺ｦ蜍慕噪・・,
    "\u30c8\u30ec\u30fc\u30ea\u30f3\u30b0": "繝医Ξ繝ｼ繝ｪ繝ｳ繧ｰ蟷・ｼ・驕主臆蛻・↓蠢懊§縺ｦ蜍慕噪・・,
    "\u6c7a\u6e08\u72b6\u6cc1": "蛻ｩ遒ｺ貂・謳榊・貂・蠑輔￠謌舌ｊ 遲・,
    "\u6700\u826f\u8cb7\u6c17\u914d\u5024": "譚ｿ縺ｮ譛濶ｯ雋ｷ豌鈴・蛟､・・SS・・,
    "\u6700\u826f\u58f2\u6c17\u914d\u5024": "譚ｿ縺ｮ譛濶ｯ螢ｲ豌鈴・蛟､・・SS・・,
    "\u30ae\u30e3\u30c3\u30d7(bp)": "(VWAP-蜑肴律邨ょ､)/蜑肴律邨ょ､*10000",
    "\u76f8\u95a2(NKY)": "驫俶氛縺ｨ譌･邨悟ｹｳ蝮・・逶ｸ髢｢菫よ焚・・1・・・・
}

PARAM_LABELS = [
    "NKY_Code","NKY_Last","NKY_ChgPct","Bias_bp","BiasSlope","GapSlope",
    "GapBanPct","NoTradeMin","TP_per_J","SL_per_J","Trail_per_J","CorrSlope",
]

PARAM_DESC = {
    "NKY_Code": "譌･邨悟ｹｳ蝮・・迚ｩ/NKY繧､繝ｳ繝・ャ繧ｯ繧ｹ遲峨・RSS繧ｳ繝ｼ繝・,
    "NKY_Last": "NKY縺ｮ迴ｾ蝨ｨ蛟､・・SS縺ｧ閾ｪ蜍墓峩譁ｰ・・,
    "NKY_ChgPct": "NKY縺ｮ鬨ｰ關ｽ邇Ⅷ%]・・SS縺ｧ閾ｪ蜍墓峩譁ｰ・・,
    "Bias_bp": "蟶ょｴ繝舌う繧｢繧ｹ[bp]・・ 鬨ｰ關ｽ邇・100・・,
    "BiasSlope": "蟶ょｴ繝舌う繧｢繧ｹ縺繰_th縺ｫ荳弱∴繧倶ｿよ焚・・100bp・・,
    "GapSlope": "繧ｮ繝｣繝・・縺繰_th縺ｫ荳弱∴繧倶ｿよ焚・・1%・・,
    "GapBanPct": "縺薙・%繧定ｶ・∴繧九ぐ繝｣繝・・縺ｧ蠖捺律蜿門ｼ慕ｦ∵ｭ｢・・AN・・,
    "NoTradeMin": "蟇・ｻ倥°繧峨％縺ｮ蛻・焚縺ｯ譁ｰ隕冗匱豕ｨ繧帝∩縺代ｋ",
    "TP_per_J": "蛻ｩ遒ｺ蟷・ｼ・驕主臆1.0縺ゅ◆繧翫・豈皮紫・・,
    "SL_per_J": "謳榊・蟷・ｼ・驕主臆1.0縺ゅ◆繧翫・豈皮紫・・,
    "Trail_per_J": "繝医Ξ繝ｼ繝ｪ繝ｳ繧ｰ蟷・ｼ・驕主臆1.0縺ゅ◆繧翫・豈皮紫・・,
    "CorrSlope": "逶ｸ髢｢ﾃ怜ｸょｴ繝舌う繧｢繧ｹ縺繰_th縺ｫ荳弱∴繧倶ｿよ焚",
}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=r"C:/AI/asagake/ASAGAKE.xlsm")
    args = ap.parse_args()

    import win32com.client  # type: ignore
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception:
        pass
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(Path(args.excel)))
        try:
            ws = wb.Worksheets("NewDashboardV2")
        except Exception:
            raise SystemExit("NewDashboardV2 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲ょ・縺ｫ build_dashboard_ui.py 繧貞ｮ溯｡後＠縺ｦ縺上□縺輔＞縲・)

        # 1) 繝代Λ繝｡繝ｼ繧ｿ隱ｬ譏趣ｼ郁｡・: 繝ｩ繝吶Ν縺ｫ繧ｳ繝｡繝ｳ繝医∬｡・: 蛟､・・        for i, label in enumerate(PARAM_LABELS, start=1):
            ws.Cells(1, i).Value = label
            try:
                cmt = ws.Cells(1, i).Comment
                if cmt is not None:
                    cmt.Delete()
            except Exception:
                pass
            try:
                ws.Cells(1, i).AddComment(PARAM_DESC.get(label, label))
            except Exception:
                pass

        # 2) 隕句・縺暦ｼ郁｡・=JP縲∬｡・=EN・画紛蜷・        for i, txt in enumerate(JP_LABELS, start=1):
            ws.Cells(4, i).Value = txt
            try:
                cmt = ws.Cells(4, i).Comment
                if cmt is not None:
                    cmt.Delete()
            except Exception:
                pass
            try:
                ws.Cells(4, i).AddComment(JP_EXPLAIN.get(txt, txt))
            except Exception:
                pass
        for i, txt in enumerate(EN_HEADERS, start=1):
            ws.Cells(5, i).Value = txt

        # 3) 繧ｹ繝・・繧ｿ繧ｹ縺ｨ繝懊ち繝ｳ縺ｮ驟咲ｽｮ繧貞承荳翫∈・磯㍾縺ｪ繧雁屓驕ｿ・・        # 繧ｹ繝・・繧ｿ繧ｹ: N1:U2
        rng = ws.Range("N1:U2")
        rng.Merge()
        rng.HorizontalAlignment = -4108
        rng.VerticalAlignment = -4108
        rng.Font.Bold = True
        rng.Font.Size = 16
        rng.Interior.Color = 15132390
        rng.Value = "IDLE"
        try:
            rng.Name = "RunStatusV2"
        except Exception:
            pass

        # 譌｢蟄倥・繧ｿ繝ｳ蜑企勁
        try:
            for shp in list(ws.Shapes):
                if shp.Name.startswith("btn_"):
                    shp.Delete()
        except Exception:
            pass

        def add_btn(name: str, caption: str, r: int, c: int, macro: str, col_span: int = 4):
            cell = ws.Cells(r, c)
            left, top = cell.Left, cell.Top
            width = ws.Range(ws.Cells(r, c), ws.Cells(r, c + col_span - 1)).Width - 4
            height = ws.Rows(r).Height - 2
            shp = ws.Shapes.AddShape(1, left, top, width, height)
            try:
                shp.Name = name
            except Exception:
                pass
            shp.TextFrame.Characters().Text = caption
            shp.TextFrame.HorizontalAlignment = -4108
            shp.TextFrame.VerticalAlignment = -4108
            shp.OnAction = macro

        # 繝懊ち繝ｳ縺ｯ陦・縺ｮ蜿ｳ蛛ｴ縺ｸ・・・朸縲・㍾縺ｪ繧峨↑縺・ｼ・        add_btn("btn_demo_start", "\u30c7\u30e2\u53d6\u5f15\u958b\u59cb", 3, 14, "AutoTraderAdvanced.StartDemoV2")  # N3
        add_btn("btn_demo_stop",  "\u30c7\u30e2\u53d6\u5f15\u505c\u6b62", 3, 18, "AutoTraderAdvanced.StopDemoV2")   # R3
        add_btn("btn_live_start", "\u672c\u756a\u53d6\u5f15\u958b\u59cb", 3, 22, "AutoTraderAdvanced.StartLiveV2") # V3
        add_btn("btn_live_stop",  "\u672c\u756a\u53d6\u5f15\u505c\u6b62", 3, 26, "AutoTraderAdvanced.StopLiveV2")  # Z3
        add_btn("btn_import",     "\u5019\u88dc\u9298\u67c4\u53d6\u8fbc", 3, 30, "AutoTraderAdvanced.ImportCandidatesV2") # AD3

        # 4) 菴楢｣∬ｪｿ謨ｴ
        ws.Rows(3).RowHeight = 28
        ws.Range("A5:AD5").EntireColumn.AutoFit()
        # 蜀ｷ蜃阪・繧､繝ｳ: 5陦後〒蝗ｺ螳夲ｼ郁ｦ句・縺怜崋螳夲ｼ・        try:
            ws.Activate()
            ws.Range("A6").Select()
            xl.ActiveWindow.FreezePanes = False
            xl.ActiveWindow.SplitRow = 5
            xl.ActiveWindow.FreezePanes = True
        except Exception:
            pass

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()


