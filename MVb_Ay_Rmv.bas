Attribute VB_Name = "MVb_Ay_Rmv"
Option Compare Binary
Option Explicit

Function Ay_XRmv_3T(A) As String()
Ay_XRmv_3T = AyMap_Sy(A, "Rmv3T")
End Function

Function Ay_XRmv_FstChr(A) As String()
Ay_XRmv_FstChr = AyMap_Sy(A, "XRmv_FstChr")
End Function

Function Ay_XRmv_FstNonLetter(A) As String()
Ay_XRmv_FstNonLetter = AyMap_Sy(A, "XRmv_FstNonLetter")
End Function

Function Ay_XRmv_LasChr(A) As String()
Ay_XRmv_LasChr = AyMap_Sy(A, "XRmv_LasChr")
End Function

Function Ay_XRmv_Pfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim U&: U = UB(A)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = XRmv_Pfx(A(J), Pfx)
Next
Ay_XRmv_Pfx = O
End Function

Function Ay_XRmv_SngQRmk(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim X, O$()
For Each X In AyNz(A)
    If Not IsSngQRmk(CStr(X)) Then Push O, X
Next
Ay_XRmv_SngQRmk = O
End Function

Function Ay_XRmv_SngQuote(A$()) As String()
Ay_XRmv_SngQuote = AyMap_Sy(A, "XRmv_SngQuote")
End Function

Function Ay_XRmv_T1(A) As String()
Dim I
For Each I In AyNz(A)
    PushI Ay_XRmv_T1, RmvT1(I)
Next
End Function


Function Ay_XRmv_TT(A$()) As String()
Ay_XRmv_TT = AyMap_Sy(A, "RmvTT")
End Function
