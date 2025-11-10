from pathlib import Path
import win32com.client as win32

WB = Path(r'C:/AI/asagake/ASAGAKE.xlsm')
AUTO = Path(r'C:/AI/asagake/excel/AutoTraderAdvanced.bas')

WATCHER_CODE = (
    'Option Explicit\r\n'
    'Public WithEvents App As Application\r\n\r\n'
    'Private Sub Class_Initialize()\r\n    Set App = Application\r\nEnd Sub\r\n\r\n'
    'Private Sub Class_Terminate()\r\n    Set App = Nothing\r\nEnd Sub\r\n\r\n'
    'Private Sub App_SheetCalculate(ByVal Sh As Object)\r\n'
    '    On Error Resume Next\r\n'
    '    Application.Run "AutoTraderAdvanced.OnDashboardCalculate", Sh\r\n'
    'End Sub\r\n'
)

def sanitize(text: str) -> str:
    text = text.replace('\ufeff','')
    out=[]
    for line in text.splitlines():
        s=line.strip().lower()
        if s.startswith('attribute vb_') or s.startswith('version ') or s in {'begin','end'}:
            continue
        out.append(line)
    return '\r\n'.join(out).rstrip()+'\r\n'

excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
wb = excel.Workbooks.Open(str(WB), ReadOnly=False)
proj = wb.VBProject
comps = proj.VBComponents

# Remove any existing AutoTraderAdvanced / cDashboardWatcher regardless of type
for name in ('AutoTraderAdvanced','cDashboardWatcher'):
    for i in range(comps.Count,0,-1):
        c = comps.Item(i)
        if c.Name.lower()==name.lower():
            comps.Remove(c)
            break

# Recreate AutoTraderAdvanced as StdModule with sanitized body
std = comps.Add(1)
std.Name = 'AutoTraderAdvanced'
mod = std.CodeModule
mod.AddFromString(sanitize(AUTO.read_text(encoding='utf-8',errors='ignore')))

# Recreate cDashboardWatcher as ClassModule with minimal code
klass = comps.Add(2)
klass.Name = 'cDashboardWatcher'
km = klass.CodeModule
km.AddFromString(WATCHER_CODE)

wb.Save()
wb.Close(True)
excel.Quit()
print('stomped')
