Attribute VB_Name = "MVb_Lin_Term"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Lin_Term."

Function Lin_TermAy(A) As String()
Dim L$, J%
L = A
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushI Lin_TermAy, XShf_Term(L)
Wend
End Function

Function XShf_T$(O)
XShf_T = XShf_Term(O)
End Function

Function XShf_X(O, X$) As Boolean
If Lin_T1(O) = X Then
    XShf_X = True
    O = RmvT1(O)
End If
End Function

Private Function XShf_Term1$(O)
Dim A$
AyAsg XBrk_Bkt(O, "["), A, XShf_Term1, O
End Function

Function XShf_Term$(O)
Dim A$
    A = LTrim(O)
If XTak_FstChr(A) = "[" Then XShf_Term = XShf_Term1(O): Exit Function
Dim P%
    P = InStr(A, " ")
If P = 0 Then
    XShf_Term = A
    O = ""
    Exit Function
End If
XShf_Term = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Private Sub Z_XShf_T()
Dim O$, OEpt$
O = " S   DFKDF SLDF  "
OEpt = "DFKDF SLDF  "
Ept = "S"
GoSub Tst
'
O = " AA BB "
Ept = "AA"
OEpt = "BB "
GoSub Tst
'
Exit Sub
Tst:
    Act = XShf_T(O)
    C
    Ass O = OEpt
    Return
End Sub


Private Sub Z()
Z_XShf_T
MVb_Lin_Term:
End Sub
