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

    win32 = win32com.client.DispatchEx("Excel.Application")
    win32.Visible = False
    win32.DisplayAlerts = False
    win32.EnableEvents = False
    win32.ScreenUpdating = False
    # 3 = msoAutomationSecurityForceDisable (protects against Auto_Open firing)
    win32.AutomationSecurity = 3
    wb = None
    try:
        wb = win32.Workbooks.Open(
            str(wb_path),
            UpdateLinks=False,
            ReadOnly=False,
            AddToMru=False,
        )
        vbproj = wb.VBProject  # type: ignore[attr-defined]
        remove_module(vbproj, "AutoTrader")
        vbcomp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule = 1
        vbcomp.Name = "AutoTrader"
        with open(bas_path, "r", encoding="utf-8") as f:
            code = normalize_module_text(f.read())
        vbcomp.CodeModule.AddFromString(code)
        wb.Save()
        print("Imported module into:", wb_path)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        win32.Quit()


def remove_module(vbproj, name: str) -> None:
    components = vbproj.VBComponents
    for index in range(components.Count, 0, -1):
        comp = components.Item(index)
        if comp.Name.lower() == name.lower():
            components.Remove(comp)
            break


def normalize_module_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text.startswith("\ufeff"):
        text = text[1:]
    lines = text.split("\n")
    if lines and lines[0].strip().lower().startswith("attribute vb_name"):
        lines = lines[1:]
    # Ensure trailing newline so AddFromString ends cleanly
    return "\r\n".join(lines).rstrip() + "\r\n"


if __name__ == "__main__":
    main()
