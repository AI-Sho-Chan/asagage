import xml.etree.ElementTree as ET
from pathlib import Path
base = Path("work/analysis_unzip/xl")
ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
root = ET.parse(base / 'worksheets' / 'sheet7.xml').getroot()
rows = root.find('main:sheetData', ns).findall('main:row', ns)
for row in rows:
    rnum = int(row.attrib['r'])
    if rnum == 6:
        for cell in row.findall('main:c', ns):
            coord = cell.attrib['r']
            if coord in ('AT6','AU6','AV6','AW6'):
                f = cell.find('main:f', ns)
                v = cell.find('main:v', ns)
                print(coord, 'formula=' + str(f.text if f is not None else None), 'value=' + str(v.text if v is not None else None))
        break
