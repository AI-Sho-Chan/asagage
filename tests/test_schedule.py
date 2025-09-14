from asagake.schedule import plan_batches

def test_plan_batches_50():
    syms = [f"{i:04d}" for i in range(200)]
    batches = plan_batches(syms, 50)
    assert len(batches) == 4
    assert all(len(b) <= 50 for b in batches)
    assert batches[0][0] == "0000" and batches[-1][-1] == "0199"
