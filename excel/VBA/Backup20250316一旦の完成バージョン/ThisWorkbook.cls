VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error Resume Next
    modAutoReload.LoadWatchlist          ' 監視リスト読み込み
    modBars.RebuildBarsAll               ' Bars を再構築
    Application.OnTime Now + TimeSerial(0, 0, 2), "modAlerts.CheckSignals"
    modAutoReload.StartWatchlistAutoReload   ' 5秒おき自動リロード開始
End Sub

