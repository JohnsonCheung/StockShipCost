Attribute VB_Name = "MVb_Lin_Term_TermN"
Option Compare Binary
Option Explicit

Function Lin_T1$(A)
Lin_T1 = Lin_TermN(A, 1)
End Function

Function Lin_T2$(A)
Lin_T2 = Lin_TermN(A, 2)
End Function

Function Lin_T3$(A)
Lin_T3 = Lin_TermN(A, 3)
End Function

Function Lin_TermN$(A, N%)
Dim L$, J%
L = A
For J = 1 To N - 1
    XShf_Term L
Next
Lin_TermN = XShf_Term(L)
End Function

Private Sub Z_Lin_TermN()
Dim N%, A$
N = 1: A = "a b c": Ept = "a": GoSub Tst
N = 2: A = "a b c": Ept = "b": GoSub Tst
N = 3: A = "a b c": Ept = "c": GoSub Tst
Exit Sub
Tst:
    Act = Lin_TermN(A, N)
    C
    Return
End Sub


Private Sub Z()
Z_Lin_TermN
MVb_Lin_Term_TermN:
End Sub
