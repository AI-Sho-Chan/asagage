Attribute VB_Name = "modNudgeRssTriggers"
Option Explicit
Public Sub NudgeRssTriggers()
    Dim ws As Worksheet: Set ws = Sheets("Bars")
    Dim t, addrs
    addrs = Array("A2", "M2", "Y2", "AK2", "AW2", "BI2", "BU2", "CG2", "CS2", "DE2", _
                  "DQ2", "EC2", "EO2", "FA2", "FM2", "FY2", "GK2", "GW2", "HI2", "HU2")
    ActiveWindow.DisplayFormulas = False
    Application.Calculation = xlCalculationAutomatic
    For Each t In addrs
        With ws.Range(t)
            If Len(.Formula2) > 0 Then .Formula2 = .Formula2   ' “¯‚¶®‚ğ‘ã“üDirty‰»
        End With
        DoEvents
    Next t
    Application.CalculateFullRebuild
End Sub


