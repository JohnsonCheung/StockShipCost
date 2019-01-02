Attribute VB_Name = "MVb_Ay_AB"
Option Compare Binary
Option Explicit

Function AyAB_XJn(A, B, Optional Sep$ = " ") As String()
Dim O$(), J&, U&
U = UB(A): If U <> UB(B) Then Stop
If U = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = A(J) & Sep & B(J)
Next
AyAB_XJn = O
End Function

Function AyAB_Dic(A, B) As Dictionary
Dim N1&, N2&
N1 = Sz(A)
N2 = Sz(B)
If N1 <> N2 Then Stop
Set AyAB_Dic = New Dictionary
Dim J&, X
For Each X In AyNz(A)
    AyAB_Dic.Add X, B(J)
    J = J + 1
Next
End Function

Function AyAB_Fmt(A, B) As String()
AyAB_Fmt = S1S2Ay_Fmt(AyabS1S2Ay(A, B))
End Function

Function AyAB_XMap_INTO(A, B, FunAB$, OInto)
Dim J&, U&, O
O = OInto
U = Min(UB(A), UB(B))
If U >= 0 Then ReDim O(U)
For J = 0 To U
    O(J) = Run(FunAB, A(J), B(J))
Next
AyAB_XMap_INTO = O
End Function

Function AyAB_XMap_SY(A, B, FunAB$) As String()
AyAB_XMap_SY = AyAB_XMap_INTO(A, B, FunAB, EmpSy)
End Function

Function AyAB_Ly_NON_EMP_B(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
AyAB_Ly_NON_EMP_B = O
End Function

Sub AyAB_XSet_SamMax(OA, OB)
Dim U1&, U2&
U1 = UB(OA)
U2 = UB(OB)
Select Case True
Case U1 > U2: ReDim Preserve OB(U1)
Case U1 < U2: ReDim Preserve OA(U2)
End Select
End Sub

Sub Ay_XAss_Eq(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
ChkAss Ay_IsEqChk(A, B, N1, N2)
End Sub
