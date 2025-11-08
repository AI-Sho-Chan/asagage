import xml.etree.ElementTree as ET
from pathlib import Path
base = Path("work/analysis_unzip/xl")
ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
# load shared strings
shared = []
ss = base / "sharedStrings.xml"
if ss.exists():
    root = ET.parse(ss).getroot()
    for si in root.findall('main:si', ns):
        text = []
        if si.find('main:t', ns) is not None:
            text.append(si.find('main:t', ns).text or '')
        else:
            for run in si.findall('main:r', ns):
                t = run.find('main:t', ns)
                if t is not None:
                    text.append(t.text or '')
        shared.append(''.join(text))

sheet = ET.parse(base / "worksheets" / "sheet7.xml").getroot()
rows = sheet.find('main:sheetData', ns).findall('main:row', ns)
# header row is row with r=5 presumably (DASH_HEADER_ROW)
headers = {}
for cell in rows[3].findall('main:c', ns):  # row index? row r="5" is 4 zero-based?? maybe 4? We'll check
    r = cell.attrib['r']
    if cell.attrib.get('t') == 's':
        v = int(cell.find('main:v', ns).text)
        value = shared[v]
    else:
        v_el = cell.find('main:v', ns)
        value = v_el.text if v_el is not None else ''
    headers[r] = value
print('header cells count', len(headers))
# get first data row
for row in rows:
    if int(row.attrib['r']) == 6:
        data_cells = {}
        for cell in row.findall('main:c', ns):
            coord = cell.attrib['r']
            t = cell.attrib.get('t')
            if t == 's':
                v = int(cell.find('main:v', ns).text)
                value = shared[v]
            else:
                v_el = cell.find('main:v', ns)
                value = v_el.text if v_el is not None else ''
            data_cells[coord] = value
        print('row6 cells', data_cells)
        break
