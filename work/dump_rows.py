import xml.etree.ElementTree as ET
from pathlib import Path
base = Path("work/analysis_unzip/xl")
ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
shared = []
ss = base / "sharedStrings.xml"
if ss.exists():
    root = ET.parse(ss).getroot()
    for si in root.findall('main:si', ns):
        text=[]
        if si.find('main:t', ns) is not None:
            text.append(si.find('main:t', ns).text or '')
        else:
            for run in si.findall('main:r', ns):
                t = run.find('main:t', ns)
                if t is not None:
                    text.append(t.text or '')
        shared.append(''.join(text))

def cell_value(cell):
    t = cell.attrib.get('t')
    v_el = cell.find('main:v', ns)
    if v_el is None:
        return ''
    val = v_el.text or ''
    if t == 's':
        idx = int(val)
        return shared[idx]
    return val

sheet = ET.parse(base / 'worksheets' / 'sheet7.xml').getroot()
rows = sheet.find('main:sheetData', ns).findall('main:row', ns)
for row in rows[:12]:
    rnum = int(row.attrib['r'])
    values = {}
    for cell in row.findall('main:c', ns):
        values[cell.attrib['r']] = cell_value(cell)
    print(rnum, values)
