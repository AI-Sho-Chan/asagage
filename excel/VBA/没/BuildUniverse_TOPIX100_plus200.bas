Attribute VB_Name = "BuildUniverse_TOPIX100_plus200"

Sub BuildUniverse_TOPIX100_plus200()
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets("Dashboard")
    Dim last As Long: last = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then MsgBox "Dashboard!A2↓にコードがありません。": Exit Sub

    '--- TOPIX100 読み込み（シートが無ければ空） ---
    Dim dictTop As Object: Set dictTop = CreateObject("Scripting.Dictionary")
    Dim wsT As Worksheet, rt As Long, lastT As Long, t As String
    On Error Resume Next: Set wsT = ThisWorkbook.Worksheets("TOPIX100"): On Error GoTo 0
    If Not wsT Is Nothing Then
        lastT = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
        For rt = 1 To lastT
            t = CleanTickerRx(CStr(wsT.Cells(rt, "A").Value))
            If Len(t) = 4 Then dictTop(t) = True
        Next
    End If

    '--- 非TOPIX100 の候補収集（AE優先。全滅ならADで代替） ---
    Dim useAE As Boolean, r As Long, code As String, nAE As Long, nAll As Long
    For r = 2 To last
        code = CleanTickerRx(CStr(wsD.Cells(r, "A").Value))
        If Len(code) = 4 And Not dictTop.Exists(code) Then
            nAll = nAll + 1
            If SafeBool(wsD.Cells(r, "AE").Value) Then nAE = nAE + 1
        End If
    Next
    useAE = (nAE > 0)

    Dim arr(), n As Long: If nAll > 0 Then ReDim arr(1 To nAll, 1 To 2)
    If nAll > 0 Then
        For r = 2 To last
            code = CleanTickerRx(CStr(wsD.Cells(r, "A").Value))
            If Len(code) = 4 And Not dictTop.Exists(code) Then
                If (useAE And SafeBool(wsD.Cells(r, "AE").Value)) Or (Not useAE) Then
                    n = n + 1
                    arr(n, 1) = code
                    arr(n, 2) = NzD(wsD.Cells(r, "AD").Value, -1E+99) ' ← #N/A 等を極小に
                End If
            End If
        Next
    End If
    If n = 0 And dictTop.Count = 0 Then
        MsgBox "候補が見つかりません。A列/AE/ADを確認してください。", vbExclamation: Exit Sub
    End If
    If n > 0 Then ReDim Preserve arr(1 To n, 1 To 2)

    ' 降順ソート
    Dim i As Long, j As Long, c As String, s As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i, 2) < arr(j, 2) Then
                c = arr(i, 1): s = arr(i, 2)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2)
                arr(j, 1) = c: arr(j, 2) = s
            End If
        Next
    Next

    ' 出力
    Dim p As String: p = Application.GetSaveAsFilename("watchlist.txt", "Text Files (*.txt),*.txt")
    If p = "False" Then Exit Sub
    Dim f As Integer: f = FreeFile: Open p For Output As #f
    Dim k As Variant

    For Each k In dictTop.Keys: Print #f, k: Next             ' TOPIX100
    For i = 1 To Application.Min(200, n): Print #f, arr(i, 1): Next  ' 残り200
    Close #f
    ThisWorkbook.Worksheets("Settings").Range("B1").Value = p
    MsgBox "watchlist.txt 作成。Settings!B1 に設定: " & p, vbInformation
End Sub

'--- ヘルパ ---
Private Function NzD(ByVal v As Variant, Optional ByVal def As Double = 0#) As Double
    On Error GoTo f
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then GoTo f
    If VarType(v) = vbString Then If Trim$(CStr(v)) = "" Then GoTo f
    NzD = CDbl(v): Exit Function
f:  NzD = def
End Function

Private Function SafeBool(ByVal v As Variant) As Boolean
    On Error GoTo f
    If IsError(v) Then GoTo f
    If VarType(v) = vbBoolean Then SafeBool = v: Exit Function
    SafeBool = (UCase$(CStr(v)) = "TRUE")
    Exit Function
f:  SafeBool = False
End Function

