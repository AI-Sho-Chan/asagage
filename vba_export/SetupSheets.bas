Attribute VB_Name = "SetupSheets"
'================= Module: SetupSheets =================
Option Explicit

' 期待ヘッダー（順序固定）
Private Const HDR As String = "code,ATR_n,TPk,SLk,J_th,dJ_th,vEMA_th,winrate,PF_eff,trades,MaxQty"

'--- 1) Allow を作り直し（シート新規＋表は空） ---
Public Sub Allow_Rebuild()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("Allow")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = "Allow"

    ' まだヘッダーは書かない（CSVをA1へ直取り込みする）
    ' ただし、表領域が無いと操作しづらいので空表を仮置き→後で作り直す
    With ws
        .Cells.Clear
        .Range("A1").value = Split(HDR, ",")  ' 一旦置く→後でCSVの1行目に置換
        .ListObjects.Add xlSrcRange, .Range("A1:K2"), , xlYes
        .ListObjects(1).name = "AllowTable"
        .Cells.Clear                           ' クリーンに戻す
    End With
End Sub

'--- 2) Allow.csv を A1 に直取り込み → ヘッダー正規化（1行のみ） ---
Public Sub Allow_LoadCsv_SingleHeader()
    Dim p As Variant, ws As Worksheet, qt As QueryTable
    p = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Allow.csv を選択")
    If VarType(p) = vbBoolean Then Exit Sub  ' False のとき

    On Error Resume Next: Set ws = Worksheets("Allow"): On Error GoTo 0
    If ws Is Nothing Then Allow_Rebuild: Set ws = Worksheets("Allow")

    Application.ScreenUpdating = False
    With ws
        .Cells.Clear
        ' ★ CSV を A1 にそのまま取り込み（TEXT;パス）? Microsoft 標準手順
        '   https://learn.microsoft.com/office/vba/api/excel.querytables.add
        Set qt = .QueryTables.Add(Connection:="TEXT;" & CStr(p), Destination:=.Range("A1"))
        With qt
            .TextFileCommaDelimiter = True
            .TextFilePlatform = 65001
            .Refresh BackgroundQuery:=False
            .Delete
        End With

        ' 取り込み後の使用範囲
        Dim lastRow As Long, lastCol As Long
        lastRow = .Cells(.rows.Count, "A").End(xlUp).row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        If lastRow < 1 Then
            Application.ScreenUpdating = True
            MsgBox "CSVが空です。", vbExclamation: Exit Sub
        End If

        ' ヘッダー正規化：1行目を期待ヘッダーへ置換し、2行目がヘッダー重複なら削除
        Dim want() As String, i As Long
        want = Split(HDR, ",")
        For i = 0 To UBound(want)
            .Cells(1, i + 1).value = want(i)
        Next i
        ' 2行目が "code" などヘッダーと同じなら削除
        If LCase$(CStr(.Cells(2, 1).value)) = "code" Then
            .rows(2).Delete
            lastRow = lastRow - 1
        End If

        ' 欠けている列（MaxQtyなど）を補う
        If lastCol < UBound(want) + 1 Then lastCol = UBound(want) + 1
        If .Cells(1, 11).value <> "MaxQty" Then .Cells(1, 11).value = "MaxQty"

        ' 既存のListObjectを破棄→今回範囲で作り直し
        On Error Resume Next
        .ListObjects(1).Unlist
        On Error GoTo 0
        .ListObjects.Add xlSrcRange, .Range(.Cells(1, 1), .Cells(lastRow, 11)), , xlYes
        .ListObjects(1).name = "AllowTable"
        .Columns.AutoFit
    End With
    Application.ScreenUpdating = True

    MsgBox "Allow.csv を取り込み（ヘッダー1行化）完了。", vbInformation
End Sub

'--- 3) Dashboard に XLOOKUP と READY を敷設（Allow 参照） ---
Public Sub Allow_ApplyDashboard()
    Dim ws As Worksheet: Set ws = Worksheets("Dashboard")
    Dim last As Long: last = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    If last < 6 Then last = 200

    ' Q:勝率 / Z:ATR_n / AA:TPk / AB:SLk / AC:J_th / AD:dJ_th / AE:vEMA_th
    ' XLOOKUP の構文: https://support.microsoft.com/office/xlookup
    ws.Range("Q6:Q" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-16],Allow!C1,Allow!C8,""""),"""")"
    ws.Range("Z6:Z" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-25],Allow!C1,Allow!C2,""""),"""")"
    ws.Range("AA6:AA" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-26],Allow!C1,Allow!C3,設定!R3C2),"""")"
    ws.Range("AB6:AB" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-27],Allow!C1,Allow!C4,設定!R4C2),"""")"
    ws.Range("AC6:AC" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-28],Allow!C1,Allow!C5,0.8),"""")"
    ws.Range("AD6:AD" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-29],Allow!C1,Allow!C6,0.02),"""")"
    ws.Range("AE6:AE" & last).FormulaR1C1 = "=IFERROR(XLOOKUP(RC[-30],Allow!C1,Allow!C7,0.02),"""")"

    ' R: READY（板累積=AK:AT は各自の列に合わせて調整）
    ws.Range("R6:R" & last).FormulaR1C1 = _
      "=IF(AND(RC[-9]<>"""",RC[-14]>0,RC[-13]>0,RC[-12]>0," & _
      "ABS(RC[-11])>=RC[12],ABS(RC[-10])>=RC[13],ABS(RC[-9])>=RC[14]," & _
      "RC[-3]>=設定!R6C2,RC[-2]<=設定!R7C2,IFERROR(RC[-3]/ABS(RC[-2]),0)>=設定!R8C2," & _
      "RC[-7]>0,SUM(RC[20]:RC[29])>=RC[-8]*設定!R10C2,IFERROR(XLOOKUP(RC[-17],Allow!C1,Allow!C1,""""),"""")<>"""")" & _
      ",""READY"","""")"

    MsgBox "Dashboard 敷設完了：Allow参照とREADY式を更新", vbInformation
End Sub
'================= End Module =================


