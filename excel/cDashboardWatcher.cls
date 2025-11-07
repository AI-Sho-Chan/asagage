VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cDashboardWatcher"
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetCalculate(ByVal Sh As Object)
    On Error Resume Next
    AutoTraderAdvanced.OnDashboardCalculate Sh
End Sub
