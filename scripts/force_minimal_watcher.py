from pathlib import Path
import win32com.client as win32

wb_path = Path(r'C:/AI/asagake/ASAGAKE.xlsm')
code = (
    'Option Explicit\r\n'
    'Public WithEvents App As Application\r\n\r\n'
    'Private Sub Class_Initialize()\r\n'
    '    Set App = Application\r\n'
    'End Sub\r\n\r\n'
    'Private Sub Class_Terminate()\r\n'
    '    Set App = Nothing\r\n'
    'End Sub\r\n\r\n'
    'Private Sub App_SheetCalculate(ByVal Sh As Object)\r\n'
    '    On Error Resume Next\r\n'
    '    Application.Run "AutoTraderAdvanced.OnDashboardCalculate", Sh\r\n'
    'End Sub\r\n'
)

excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
wb = excel.Workbooks.Open(str(wb_path), ReadOnly=False)
proj = wb.VBProject
components = proj.VBComponents
name = 'cDashboardWatcher'
# Remove existing
for idx in range(components.Count, 0, -1):
    comp = components.Item(idx)
    if comp.Name.lower() == name.lower():
        components.Remove(comp)
        break
# Add fresh class and inject minimal code (no VERSION/Attribute headers)
vbext_ct_ClassModule = 2
comp = components.Add(vbext_ct_ClassModule)
comp.Name = name
mod = comp.CodeModule
if mod.CountOfLines > 0:
    mod.DeleteLines(1, mod.CountOfLines)
mod.AddFromString(code)
wb.Save()
wb.Close(True)
excel.Quit()
