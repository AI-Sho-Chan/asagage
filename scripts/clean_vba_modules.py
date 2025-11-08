import argparse
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=r"C:/AI/asagake/ASAGAKE.xlsm")
    args = ap.parse_args()

    import win32com.client  # type: ignore
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(Path(args.excel)))
        vb = wb.VBProject
        # vbext_ct_StdModule = 1
        to_remove = []
        for i in range(1, vb.VBComponents.Count + 1):
            comp = vb.VBComponents.Item(i)
            name = comp.Name
            try:
                ctype = comp.Type
            except Exception:
                ctype = None
            if ctype == 1 and name not in ("AutoTraderAdvanced",):
                # 標準モジュールで不要なものは削除（Module1.. 等）
                to_remove.append(name)
        for name in to_remove:
            try:
                vb.VBComponents.Remove(vb.VBComponents(name))
            except Exception:
                pass
        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()

