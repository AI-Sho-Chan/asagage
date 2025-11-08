from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

wb_path = Path("C:/AI/asagake/SHINSOKU.xlsm")
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(str(wb_path))
ws = wb.Worksheets("NewDashboard")
try:
    ws.Range("Z6").Formula = '=SUMPRODUCT(IFERROR(ABS($J$6:$J$406),0),IF($X$6:$X$406=1,1,0))'
    print("Z6 SUMABSJ=", ws.Range("Z6").Value)
    ws.Range("Z7").Formula = '=COUNTIF($X$6:$X$406,1)'
    print("Z7 CNTSEL=", ws.Range("Z7").Value)
    ws.Range("Z8").Formula = '=IF(S6<>"",S6,IF(P6<>"",P6,N6))'
    print("Z8 PRICE=", ws.Range("Z8").Value)
    ws.Range("Z9").Formula = '=ABS($B$16)/10000'
    print("Z9 SLIP=", ws.Range("Z9").Value)
    ws.Range("Z10").Formula = '=IFERROR(Z8*(1+Z9*IF(J6<0,1,-1)),Z8)'
    print("Z10 WORST=", ws.Range("Z10").Value)
    ws.Range("Z11").Formula = '=MAX(1,$B$15)'
    print("Z11 STEP=", ws.Range("Z11").Value)
    ws.Range("Z12").Formula = '=IF(Z6>0,$B$14*ABS(J6)/Z6,IF(Z7>0,$B$14/Z7,0))'
    print("Z12 ALLOC=", ws.Range("Z12").Value)
    ws.Range("Z13").Formula = '=IF(AND(ISNUMBER(Z10),Z10>0),Z12/Z10,IF(AND(ISNUMBER(Z8),Z8>0),Z12/Z8,0))'
    print("Z13 QBASE=", ws.Range("Z13").Value)
    ws.Range("Z14").Formula = '=INT(Z13/Z11)*Z11'
    print("Z14 QRAW=", ws.Range("Z14").Value)
    ws.Range("Z15").Formula = '=IF(Z14>0,Z14,Z11)'
    print("Z15 Q=", ws.Range("Z15").Value)
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
