Attribute VB_Name = "MXls_Z_Lo_Fmt_Tit"
Option Compare Binary
Option Explicit

Sub Lo_XSet_Tit(A As ListObject, TitLy$())
Dim Sq(), R As Range
    Sq = WSq(TitLy, Lo_Fny(A)): If Sz(Sq) = 0 Then Exit Sub
    Set R = WRg(Sq, A)
Sq_XPut_AtCell Sq, R
XMsg R
Rg_XBdr_Inner R
Rg_XBdr_Around R
End Sub

Private Sub XMsg(A As Range)
Dim J%
For J = 1 To A.Rows.Count
    XMsg_H RgR(A, J)
Next
For J = 1 To A.Columns.Count
    XMsg_V RgC(A, J)
Next
End Sub

Private Sub XMsg_H(A As Range)
A.Application.DisplayAlerts = False
Dim J%, C1%, C2%, V, LasV
LasV = RgRC(A, 1, 1).Value
C1 = 1
For J = 2 To A.Columns.Count
    V = RgRC(A, 1, J).Value
    If V <> LasV Then
        C2 = J - 1
        If Not IsEmpty(LasV) Then
            RgRCC(A, 1, C1, C2).MergeCells = True
        End If
        C1 = J
        LasV = V
    End If
Next
A.Application.DisplayAlerts = True
End Sub

Private Sub XMsg_V(A As Range)
Dim J%
For J = A.Rows.Count To 2 Step -1
    Cell_XMge_Above RgRC(A, J, 1)
Next
End Sub

Private Function WRg(Sq(), Lo As ListObject) As Range
Dim At As Range, NR%, R%
R = Lo.DataBodyRange.Row
NR = UBound(Sq, 1)
Set WRg = CellSq_XReSz(RgRC(Lo.DataBodyRange, 0 - NR, 1), Lo)
End Function

Private Function WSq(TitLy$(), Fny0) As Variant()
Dim Col()
    Dim F, Tit$
    For Each F In Ssl_Sy(Fny0)
        Tit = Ay_FstRmvT1(TitLy, F)
        If Tit = "" Then
            PushI Col, Ap_Sy(F)
        Else
            PushI Col, AyTrim(SplitVBar(Tit))
        End If
    Next
WSq = Sq_Transpose(Dry_Sq(Col))
End Function


Private Sub Z_WSq()
Dim TitLy$(), Fny0$
'----
Dim A$()
'TitLy
    Erase A
    Push A, "A A1 | A2 11 "
    Push A, "B B1 | B2 | B3"
    Push A, "C C1"
    Push A, "E E1"
    TitLy = A
Fny0 = "A B C D E"
Ept = New_Sq(3, 5)
    Sq_Set_Row Ept, 1, Ssl_Sy("A1 B1 C1 D E1")
    Sq_Set_Row Ept, 2, Array("A2 11", "B2")
    Sq_Set_Row Ept, 3, Array(Empty, "B3")
GoSub Tst
Exit Sub
'---
'TitLy
    Erase A
    Push A, "ksdf | skdfj  |skldf jf"
    Push A, "skldf|sdkfl|lskdf|slkdfj"
    Push A, "askdfj|sldkf"
    Push A, "fskldf"
Sq_Brw WSq(A, "")

Exit Sub
Tst:
    Act = WSq(TitLy, Fny0)
    Ass Sq_IsEq(Act, Ept)
    Return
End Sub

Private Sub Z()
Z_WSq
MXls_Z_Lo_Fmt_Tit:
End Sub
