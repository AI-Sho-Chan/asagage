import sys
from pathlib import Path


def main():
    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover
        print("win32com is required. Install pywin32.", exc)
        sys.exit(1)

    if len(sys.argv) < 3:
        print(
            "Usage: python scripts/excel_install_macros.py "
            "C:/AI/asagake/SHINSOKU.xlsm excel/AutoTraderAdvanced.bas [excel/cDashboardWatcher.cls ...]"
        )
        sys.exit(2)

    wb_path = Path(sys.argv[1])
    if not wb_path.exists():
        print("Workbook not found:", wb_path)
        sys.exit(3)
    module_paths = [Path(p) for p in sys.argv[2:]]
    for mod in module_paths:
        if not mod.exists():
            print("Module file not found:", mod)
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
        for module_path in module_paths:
            install_module(vbproj, module_path)
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


def extract_module_name(text: str, default: str) -> str:
    for line in text.splitlines():
        line = line.strip()
        if line.lower().startswith("attribute vb_name"):
            start = line.find('"')
            end = line.rfind('"')
            if start >= 0 and end > start:
                return line[start + 1 : end]
    return default


def install_module(vbproj, module_path: Path) -> None:
    raw_text = module_path.read_text(encoding="utf-8")
    module_name = extract_module_name(raw_text, module_path.stem)
    code = normalize_module_text(raw_text)
    remove_module(vbproj, module_name)
    ext = module_path.suffix.lower()
    if ext == ".cls":
        comp_type = 2  # vbext_ct_ClassModule
    else:
        comp_type = 1  # standard module
    vbcomp = vbproj.VBComponents.Add(comp_type)
    vbcomp.Name = module_name
    vbcomp.CodeModule.AddFromString(code)


if __name__ == "__main__":
    main()
