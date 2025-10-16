
Attribute VB_Name = "W_UDFs_Min"
Option Explicit
Private Function EV(ByVal s As String) As Variant
    On Error GoTo e: EV = Application.Evaluate(s): Exit Function
e: EV = CVErr(xlErrNA)
End Function
Private Function TickSize(ByVal px As Double) As Double
    Select Case px
        Case Is < 1000#: TickSize = 1#
        Case Is < 5000#: TickSize = 1#
        Case Is < 10000#: TickSize = 1#
        Case Else: TickSize = 1#
    End Select
End Function
Public Function W_BOOK_IMBAL(ByVal code As Variant, ByVal depth As Long) As Double
    Dim b As Variant, a As Variant
    b = EV("RssMarket(""" & Format(code, "0") & """,""最良買気配数量"")")
    a = EV("RssMarket(""" & Format(code, "0") & """,""最良売気配数量"")")
    If IsError(b) Or IsError(a) Or Not IsNumeric(b) Or Not IsNumeric(a) Then
        W_BOOK_IMBAL = 0#
    ElseIf (CDbl(b) + CDbl(a)) = 0# Then
        W_BOOK_IMBAL = 0#
    Else
        W_BOOK_IMBAL = (CDbl(a) - CDbl(b)) / (CDbl(a) + CDbl(b))
    End If
End Function
Public Function W_ENTRY_SLIP_CAP(ByVal code As Variant, ByVal side As String, _
    ByVal px As Double, ByVal kEntry As Double, ByVal kExit As Double, _
    ByVal depth As Long, ByVal feePerShare As Double) As Double
    Dim ts As Double: ts = TickSize(px)
    W_ENTRY_SLIP_CAP = Application.Max(0#, ts * kEntry) + feePerShare
End Function
Public Function W_EXIT_SLIP_CAP(ByVal code As Variant, ByVal side As String, _
    ByVal px As Double, ByVal qty As Double, ByVal depth As Long, _
    ByVal feePerShare As Double) As Double
    Dim ts As Double, crowd As Double: ts = TickSize(px)
    If qty < 1# Then qty = 1#
    If qty < 10# Then
        crowd = 1#
    Else
        crowd = 1# + WorksheetFunction.Min(5#, WorksheetFunction.Log10(qty))
    End If
    W_EXIT_SLIP_CAP = ts * crowd + feePerShare
End Function
Public Function W_QTY_BY_BUDGET_CLIPPED(ByVal px As Double, ByVal kEntry As Double, _
    ByVal kExit As Double, ByVal code As Variant, ByVal side As String, _
    ByVal depth As Long) As Double
    Dim budget As Double, q As Double: budget = 500000#
    If px <= 0# Then W_QTY_BY_BUDGET_CLIPPED = 0#: Exit Function
    On Error Resume Next
    q = Application.WorksheetFunction.Floor_Precise(budget / px, 100#)
    On Error GoTo 0
    If q <= 0# Then q = Int(budget / px / 100#) * 100#
    If q < 100# Then q = 100#
    W_QTY_BY_BUDGET_CLIPPED = q
End Function
