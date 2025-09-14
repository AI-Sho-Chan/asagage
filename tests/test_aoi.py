import numpy as np
from asagake.aoi import compute_aoi_series, summarize_aoi

def test_compute_and_summarize():
    aoi = compute_aoi_series([100, 120, 130], [50, 80, 130])
    assert len(aoi) == 3
    assert abs(aoi[0] - ((100-50)/(100+50))) < 1e-9
    m = summarize_aoi(aoi)
    assert m.n == 3
    assert m.latest == 0.0
    assert abs(m.sigma - np.std(aoi)) < 1e-12
