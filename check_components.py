from pathlib import Path
import win32com.client as win32
path = Path(r"C:/AI/asagake/ASAGAKE.xlsm")
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
wb = excel.Workbooks.Open(str(path), ReadOnly=True)
components = wb.VBProject.VBComponents
info = [(comp.Name, int(getattr(comp, "Type", -1))) for comp in components]
print(info)
wb.Close(False)
excel.Quit()
