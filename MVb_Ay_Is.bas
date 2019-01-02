Attribute VB_Name = "MVb_Ay_Is"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Ay_Is."

Function Ay_IsAllEleEq(A) As Boolean
If Sz(A) = 0 Then Ay_IsAllEleEq = True: Exit Function
Dim J&
For J = 1 To UB(A)
    If A(0) <> A(J) Then Exit Function
Next
Ay_IsAllEleEq = True
End Function

Function Ay_IsAllEleHasVal(A) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
    If IsEmp(I) Then Exit Function
Next
Ay_IsAllEleHasVal = True
End Function

Function Ay_IsAllEq(A) As Boolean
If Sz(A) <= 1 Then Ay_IsAllEq = True: Exit Function
Dim A0, J&
A0 = A(0)
For J = 2 To UB(A)
    If A0 <> A(0) Then Exit Function
Next
Ay_IsAllEq = True
End Function

Function Ay_IsAllStr(A) As Boolean
Const CSub$ = CMod & "Ay_IsAllStr"
If Not IsArray(A) Then XThw CSub, "A is not array", "TypeName(A)", TypeName(A)
If IsSy(A) Then Ay_IsAllStr = True: Exit Function
Dim I
For Each I In AyNz(A)
    If Not IsStr(I) Then Exit Function
Next
Ay_IsAllStr = True
End Function

Function Ay_IsEq(A, B) As Boolean
Const CSub$ = CMod & "Ay_IsEq"
If VarType(A) <> VarType(B) Then Exit Function
If Not IsArray(A) Then XThw CSub, "[A] is not array", "TypeName(A)", TypeName(A)
If Not Ay_IsEqSz(A, B) Then Exit Function
Dim J&, X
For Each X In AyNz(A)
    If Not IsEq(X, B(J)) Then Exit Function
    J = J + 1
Next
Ay_IsEq = True
End Function

Sub Ay_IsEq_XAss(A, B, Optional Fun$ = "Ay_IsEq_XAss")
IsEqTy_XAss A, B
IsAy_XAss A
Ay_IsEqSz_XAss A, B, Fun
If Sz(A) = 0 Then Exit Sub
Dim J&, X
For Each X In A
    IsEq_XAss X, B(J), Fun
    J = J + 1
Next
End Sub

Function Ay_IsEqSz(A, B) As Boolean
Ay_IsEqSz = Sz(A) = Sz(B)
End Function

Sub Ay_IsEqSz_XAss(A, B, Optional Fun$ = "Ay_IsEqSz_XAss")
If Not Ay_IsEqSz(A, B) Then XDmp_Lin_Stop Fun, "Siz is dif", "A-Sz B-Sz", Sz(A), Sz(B)
End Sub

Function Ay_IsLinesAy(A) As Boolean
If Not Ay_IsAllStr(A) Then Exit Function
Dim L
For Each L In AyNz(A)
    If HasCrLf(L) Then Ay_IsLinesAy = True: Exit Function
Next
End Function

Function Ay_IsSam(A, B) As Boolean
Ay_IsSam = Dic_IsEq(AyCntDic(A), AyCntDic(B))
End Function

Sub IsAy_XAss(A, Optional Fun$ = "IsAy_XAss")
If Not IsArray(A) Then
    MsgNyAp_XDmp_Lin_Stop "A is not array", "A-Ty", TypeName(A)
End If
End Sub

Sub IsEqTy_XAss(A, B)
If VarType(A) <> VarType(B) Then
    MsgNyAp_XDmp_Lin_Stop "A & B are diff Type", "A-Ty B-Ty", TypeName(A), TypeName(B)
End If
End Sub

Private Sub ZZ()
Dim A As Variant
Ay_IsAllEleEq A
Ay_IsAllEleHasVal A
Ay_IsAllEq A
Ay_IsAllStr A
Ay_IsEq A, A
Ay_IsEq_XAss A, A
Ay_IsEqSz A, A
Ay_IsEqSz_XAss A, A
Ay_IsLinesAy A
Ay_IsSam A, A
IsAy_XAss A
IsEqTy_XAss A, A
End Sub

Private Sub Z()
End Sub
