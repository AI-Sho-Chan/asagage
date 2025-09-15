Attribute VB_Name = "modRepair"
Option Explicit

Public Sub RepairAndRebuild()
    On Error GoTo EH
    FixCalcAndRefresh       '計算設定を正常化
    RebuildBarsAll          'Bars を再配置・ヘッダー復元
    NudgeRssTriggers        'RssChart 式を書き込み
    Exit Sub
EH:
    MsgBox "RepairAndRebuild error: " & Err.Description
End Sub

