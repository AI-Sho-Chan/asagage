import argparse
import shutil
from pathlib import Path


def backup(path: Path) -> Path:
    path = path.resolve()
    ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_backup_{ts}{path.suffix}")
    shutil.copy2(path, bak)
    return bak


def import_module(excel_path: Path, bas_path: Path) -> None:
    import win32com.client  # type: ignore
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityLow
    try:
        wb = excel.Workbooks.Open(str(excel_path))
        vbproj = wb.VBProject
        # Replace if exists then import from file to avoid encoding/syntax mishaps
        try:
            vbcomp = vbproj.VBComponents("AutoTraderAdvanced")
            vbproj.VBComponents.Remove(vbcomp)
        except Exception:
            pass
        vbcomp = vbproj.VBComponents.Import(str(bas_path))
        try:
            vbcomp.Name = "AutoTraderAdvanced"
        except Exception:
            pass
        # remove any leftover standard modules besides AutoTraderAdvanced
        to_remove = []
        for idx in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents.Item(idx)
            try:
                ctype = comp.Type
            except Exception:
                ctype = None
            if ctype == 1 and comp.Name not in ("AutoTraderAdvanced",):
                to_remove.append(comp.Name)
        for name in to_remove:
            try:
                vbproj.VBComponents.Remove(vbproj.VBComponents(name))
            except Exception:
                pass
        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=r"C:/AI/asagake/SHINSOKU.xlsm")
    ap.add_argument("--bas", default=r"excel/AutoTraderAdvanced.bas")
    args = ap.parse_args()

    xlsm = Path(args.excel)
    bas = Path(args.bas)
    backup(xlsm)
    import_module(xlsm, bas)


if __name__ == "__main__":
    main()
