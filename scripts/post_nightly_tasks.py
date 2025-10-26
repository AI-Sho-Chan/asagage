import argparse
import subprocess
from pathlib import Path


def parse_status(path: Path) -> dict:
    d = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            d[k.strip()] = v.strip()
    return d


def run(cmd: list[str]) -> None:
    print("[run]", " ".join(cmd))
    subprocess.run(cmd, check=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", default="logs/nightly_status.txt")
    ap.add_argument("--excel", default=r"C:/AI/asagake/SHINSOKU.xlsm")
    args = ap.parse_args()

    st = parse_status(Path(args.status))
    if st.get("state") != "success":
        print("Nightly state is not success; skipping post tasks")
        return

    cand_path = st.get("candidates_path") or "output/excel/candidates_nextday.csv"
    date_tag = st.get("date_tag") or ""

    # Export Excel logs (read-only)
    run(["python", "scripts/export_excel_logs.py", "--excel", args.excel, "--outdir", f"output/trade_logs/{date_tag}"])
    # Make size plan
    out = f"output/excel/size_plan/size_plan_{date_tag}.csv"
    run(["python", "scripts/make_size_plan.py", "--candidates", cand_path, "--out", out])


if __name__ == "__main__":
    main()

