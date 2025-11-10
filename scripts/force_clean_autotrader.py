from pathlib import Path
import re
import win32com.client as win32

SRC = Path(r"C:/AI/asagake/excel/AutoTraderAdvanced.bas")
TARGET = Path(r"C:/AI/asagake/ASAGAKE.xlsm")

raw = SRC.read_text(encoding="utf-8", errors="ignore")
# Strip BOM and any Attribute/Version lines to avoid VBIDE parsing quirks
text = raw.replace("\ufeff", "")
lines = []
for line in text.splitlines():
    s = line.strip().lower()
    if s.startswith("attribute vb_") or s.startswith("version ") or s in {"begin", "end"}:
        continue
    lines.append(line)
clean = "\r\n".join(lines) + "\r\n"

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
wb = excel.Workbooks.Open(str(TARGET), ReadOnly=False)
proj = wb.VBProject
components = proj.VBComponents
# Remove existing module
for idx in range(components.Count, 0, -1):
    comp = components.Item(idx)
    if comp.Name.lower() == "autotraderadvanced":
        components.Remove(comp)
        break
# Add standard module and inject clean code
vbext_ct_StdModule = 1
comp = components.Add(vbext_ct_StdModule)
comp.Name = "AutoTraderAdvanced"
mod = comp.CodeModule
if mod.CountOfLines > 0:
    mod.DeleteLines(1, mod.CountOfLines)
mod.AddFromString(clean)
wb.Save()
wb.Close(True)
excel.Quit()
