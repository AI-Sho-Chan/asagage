Attribute VB_Name = "modBarsSetup"
Option Explicit

Private Function HeaderArray() As Variant
    HeaderArray = Array("銘柄名称", "市場名", "足種", "日付", "時刻", "始値", "高値", "安値", "終値", "出来高")
End Function

Public Sub RebuildBarsAll()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsB As Worksheet: Set wsB = wb.Sheets("Bars")
    Dim wsD As Worksheet: Set wsD = wb.Sheets("Dashboard")

    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.rows.Count, "A").End(xlUp).Row
    Dim blocks As Long
    blocks = Application.Min(20, Application.Max(0, lastRow - 1))
    If blocks = 0 Then wsB.Cells.Clear: GoTo Tidy

    'ヘッダーと領域を一旦クリア
    wsB.Range(wsB.Cells(2, 1), wsB.Cells(22, 12 * blocks + 1)).Clear

    '各ブロックにヘッダーを書き込み
    Dim h, r As Long, k As Long, sc As Long
    h = HeaderArray()
    For r = 1 To blocks
        sc = 2 + (r - 1) * 12
        For k = 0 To UBound(h)
            wsB.Cells(2, sc + k).Value = h(k)
        Next k
    Next r

Tidy:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "RebuildBarsAll error: " & Err.Description
End Sub

