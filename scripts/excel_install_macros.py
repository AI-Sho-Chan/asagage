import sys
from pathlib import Path


def main():
    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover
        print("win32com is required. Install pywin32.", exc)
        sys.exit(1)

    if len(sys.argv) < 3:
        print("Usage: python scripts/excel_install_macros.py C:/AI/asagake/SHINSOKU.xlsm excel/vba/AutoTrader.bas")
        sys.exit(2)

    wb_path = Path(sys.argv[1])
    bas_path = Path(sys.argv[2])
    if not wb_path.exists():
        print("Workbook not found:", wb_path)
        sys.exit(3)
    if not bas_path.exists():
        print("BAS file not found:", bas_path)
        sys.exit(4)

    win32 = win32com.client.Dispatch("Excel.Application")
    win32.Visible = False
    try:
        wb = win32.Workbooks.Open(str(wb_path))
        vbproj = wb.VBProject
        remove_module(vbproj, "AutoTrader")
        vbcomp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule = 1
        vbcomp.Name = "AutoTrader"
        with open(bas_path, "r", encoding="utf-8") as f:
            code = f.read()
        vbcomp.CodeModule.AddFromString(code)
        wb.Save()
        print("Imported module into:", wb_path)
    finally:
        wb.Close(SaveChanges=True)
        win32.Quit()


def remove_module(vbproj, name: str) -> None:
    components = vbproj.VBComponents
    for index in range(components.Count, 0, -1):
        comp = components.Item(index)
        if comp.Name.lower() == name.lower():
            components.Remove(comp)
            break


if __name__ == "__main__":
    main()
