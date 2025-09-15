Attribute VB_Name = "modUDF"
Option Explicit

' B2:K22 �� 1 �u���b�N�������ɂƂ�,���� AVWAP�i�o�������d�I�l���ρj��Ԃ�
Public Function AVWAP_FromBlock(block As Range) As Variant
    On Error GoTo EH
    Dim r As Long, v As Double, c As Double
    Dim sumV As Double, sumPxV As Double

    For r = 3 To block.rows.Count
        v = Nz(block.Cells(r, 9).Value)     '�o�����iB:K �� 9 ��ځj
        c = Nz(block.Cells(r, 8).Value)     '�I�l�i8 ��ځj
        If v > 0 And c > 0 Then
            sumV = sumV + v
            sumPxV = sumPxV + c * v
        End If
    Next
    If sumV = 0 Then AVWAP_FromBlock = CVErr(xlErrNA) Else AVWAP_FromBlock = sumPxV / sumV
    Exit Function
EH:
    AVWAP_FromBlock = CVErr(xlErrNA)
End Function

' ���� 5 �{�� ATR�iTR �̒P�����ρj
Public Function ATR5_FromBlock(block As Range) As Variant
    On Error GoTo EH
    Dim i As Long, r As Long, cnt As Long
    Dim prevClose As Double, hi As Double, lo As Double, cl As Double
    Dim tr As Double, sumTR As Double

    ' ��������ŏI�I�l�s��T��
    For i = block.rows.Count To 3 Step -1
        If Nz(block.Cells(i, 8).Value) > 0 Then
            prevClose = Nz(block.Cells(i, 8).Value)
            Exit For
        End If
    Next i
    If i = 2 Then ATR5_FromBlock = CVErr(xlErrNA): Exit Function

    For r = i To 3 Step -1
        hi = Nz(block.Cells(r, 6).Value)
        lo = Nz(block.Cells(r, 7).Value)
        cl = Nz(block.Cells(r, 8).Value)
        tr = Application.WorksheetFunction.Max(hi - lo, Abs(hi - prevClose), Abs(lo - prevClose))
        sumTR = sumTR + tr
        cnt = cnt + 1
        If cnt = 5 Then Exit For
        prevClose = cl
    Next r

    If cnt = 0 Then ATR5_FromBlock = CVErr(xlErrNA) Else ATR5_FromBlock = sumTR / cnt
    Exit Function
EH:
    ATR5_FromBlock = CVErr(xlErrNA)
End Function

' Dashboard �̍s�ԍ� �� �u���b�N�Q�ƁiByRow�j
Public Function AVWAP_ByRow(rw As Long) As Variant
    Dim sc As Long: sc = 2 + (rw - 2) * 12
    With ThisWorkbook.Sheets("Bars")
        AVWAP_ByRow = AVWAP_FromBlock(.Range(.Cells(2, sc), .Cells(22, sc + 9)))
    End With
End Function

Public Function ATR5_ByRow(rw As Long) As Variant
    Dim sc As Long: sc = 2 + (rw - 2) * 12
    With ThisWorkbook.Sheets("Bars")
        ATR5_ByRow = ATR5_FromBlock(.Range(.Cells(2, sc), .Cells(22, sc + 9)))
    End With
End Function

' �R�[�h�iA��j����v�Z�iByCode�j
Public Function AVWAP_ByCode(code As Variant) As Variant
    Dim rw As Variant
    rw = Application.Match(code, ThisWorkbook.Sheets("Dashboard").Range("A:A"), 0)
    If IsError(rw) Then AVWAP_ByCode = CVErr(xlErrNA) Else AVWAP_ByCode = AVWAP_ByRow(CLng(rw))
End Function

Public Function ATR5_ByCode(code As Variant) As Variant
    Dim rw As Variant
    rw = Application.Match(code, ThisWorkbook.Sheets("Dashboard").Range("A:A"), 0)
    If IsError(rw) Then ATR5_ByCode = CVErr(xlErrNA) Else ATR5_ByCode = ATR5_ByRow(CLng(rw))
End Function

