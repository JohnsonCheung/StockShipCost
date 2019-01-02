Attribute VB_Name = "MVb_Lin_Term_NTermRst"
Option Compare Binary
Option Explicit
Function Lin_TermRst(A) As String()
Lin_TermRst = Lin_1TRst(A)
End Function
Function Lin_1TRst(A) As String()
Lin_1TRst = LinN_NTermRst(A, 1)
End Function

Function Lin_2TRst(A) As String()
Lin_2TRst = LinN_NTermRst(A, 2)
End Function

Function Lin_3TRst(A) As String()
Lin_3TRst = LinN_NTermRst(A, 3)
End Function

Function Lin_4TRst(A) As String()
Lin_4TRst = LinN_NTermRst(A, 4)
End Function

Function LinN_NTermRst(A, N%) As String()
Dim L$, J%
L = A
For J = 1 To N
    PushI LinN_NTermRst, XShf_T(L)
Next
PushI LinN_NTermRst, L
End Function

Private Sub Z_LinN_NTermRst()
Dim A$
A = "  [ksldfj ]":  Ept = "ksldfj ": GoSub Tst
A = "  [ ksldfj ]": Ept = " ksldf ": GoSub Tst
A = "  [ksldfj]":  Ept = "ksldf": GoSub Tst
Exit Sub
Tst:
    Act = Lin_T1(A)
    C
    Return
End Sub


Private Sub Z()
Z_LinN_NTermRst
MVb_Lin_Term_NTermRst:
End Sub
