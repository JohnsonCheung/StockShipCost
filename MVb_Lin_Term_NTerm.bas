Attribute VB_Name = "MVb_Lin_Term_NTerm"
Option Compare Binary
Option Explicit

Function Lin2T(A) As String()
Lin2T = LinNTerm(A, 2)
End Function

Function Lin3T(A) As String()
Lin3T = LinNTerm(A, 3)
End Function

Function LinNTerm(A, N%) As String()
Dim J%, L$
L = A
For J = 1 To N
    PushI LinNTerm, XShf_T(L)
Next
End Function

Function Lin_TT(ByVal A) As String()
Dim T1$, T2$
T1 = XShf_Term(A)
T2 = XShf_Term(A)
Lin_TT = Ap_Sy(T1, T2)
End Function
