Attribute VB_Name = "modRepair"
Option Explicit

Public Sub RepairAndRebuild()
    On Error GoTo EH
    FixCalcAndRefresh       '�v�Z�ݒ�𐳏퉻
    RebuildBarsAll          'Bars ���Ĕz�u�E�w�b�_�[����
    NudgeRssTriggers        'RssChart ������������
    Exit Sub
EH:
    MsgBox "RepairAndRebuild error: " & Err.Description
End Sub

