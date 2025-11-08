import xml.etree.ElementTree as ET
from pathlib import Path
base = Path(r'work/diagnose_unzip/xl')
ns = {'main':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
shared=[]
ss=base/'sharedStrings.xml'
if ss.exists():
    root=ET.parse(ss).getroot()
    for si in root.findall('main:si', ns):
        text=[]
        if si.find('main:t', ns) is not None:
            text.append(si.find('main:t', ns).text or '')
        else:
            for run in si.findall('main:r', ns):
                t=run.find('main:t', ns)
                if t is not None:
                    text.append(t.text or '')
        shared.append(''.join(text))

sheet=ET.parse(base/'worksheets'/'sheet7.xml').getroot()
rows=sheet.find('main:sheetData', ns).findall('main:row', ns)
# find column index for J_th column (AD row 5) to check formulas row6
for row in rows:
    if row.attrib['r']=='5':
        headers={}
        for c in row.findall('main:c', ns):
            ref=c.attrib['r']
            v=c.find('main:v', ns)
            val=''
            if v is not None:
                if c.attrib.get('t')=='s':
                    val=shared[int(v.text)]
                else:
                    val=v.text
            headers[ref]=val
        print('headers', headers.get('AD5'))
    if row.attrib['r']=='6':
        for c in row.findall('main:c', ns):
            if c.attrib['r']=='AD6':
                f=c.find('main:f', ns)
                v=c.find('main:v', ns)
                fv=f.text if f is not None else None
                vv=v.text if v is not None else None
                print('AD6 formula', fv, 'value', vv)
            if c.attrib['r']=='AT6':
                f=c.find('main:f', ns); v=c.find('main:v', ns)
                print('AT6 formula', f.text if f is not None else None, 'value', v.text if v is not None else None)
            if c.attrib['r']=='AU6':
                f=c.find('main:f', ns); v=c.find('main:v', ns)
                print('AU6 formula', f.text if f is not None else None, 'value', v.text if v is not None else None)
        break
