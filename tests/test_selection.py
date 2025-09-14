from asagake.screener import select_symbols, SelectionParams

def test_selection_filters_abs_sigma_and_samples():
    # 7203: AOI大&安定、9984: 変動大で除外、6758: 負側AOIで通過、A: サンプル不足で除外
    inputs = {
        "7203": ([120,130,140,150,155,158,160,162],[50,55,60,65,66,66,67,67]),
        "9984": ([100,200,100,200,100,200,100,200],[200,100,200,100,200,100,200,100]),
        "6758": ([50,50,55,60,62,65,66,68],[120,120,110,100,98,95,94,180]),
        "A":    ([100,110],[90,95]),  # 2点 → 除外
    }
    sel = select_symbols(inputs, SelectionParams(min_abs_aoi=0.4, max_sigma=0.2, top_k=10, min_samples=8))
    assert "7203" in sel and "6758" in sel and "9984" not in sel and "A" not in sel
