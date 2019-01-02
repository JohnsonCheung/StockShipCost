Attribute VB_Name = "MVb_Ay_BoolAy"
Option Compare Binary
Option Explicit

Enum eBoolOp
    eOpEQ = 1
    eOpNE = 2
    eOpAND = 3
    eOpOR = 4
End Enum
Enum eEqNeOp
    eOpEQ = eBoolOp.eOpEQ
    eOpNE = eBoolOp.eOpNE
End Enum
Enum eAndOrOp
    eOpAND = eBoolOp.eOpAND
    eOpOR = eBoolOp.eOpOR
End Enum

Function BoolAy_XAnd(A() As Boolean) As Boolean
BoolAy_XAnd = BoolAy_IsAllTrue(A)
End Function

Function BoolAy_IsAllFalse(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then Exit Function
Next
BoolAy_IsAllFalse = True
End Function

Function BoolAy_IsAllTrue(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
BoolAy_IsAllTrue = True
End Function

Function BoolAy_IsSomTrue(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then BoolAy_IsSomTrue = True: Exit Function
Next
End Function

Function BoolAy_XOr(A() As Boolean) As Boolean
BoolAy_XOr = BoolAy_IsSomTrue(A)
End Function

Function BoolAy_IsSomFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then BoolAy_IsSomFalse = True: Exit Function
Next
End Function


Function BoolOpStr_BoolOp(BoolOpStr) As eBoolOp
Dim O As eBoolOp
Select Case UCase(BoolOpStr)
Case "AND": O = eBoolOp.eOpAND
Case "OR": O = eBoolOp.eOpOR
Case "EQ": O = eBoolOp.eOpEQ
Case "NE": O = eBoolOp.eOpNE
Case Else: Stop
End Select
BoolOpStr_BoolOp = O
End Function

Function BoolOpStr_IsAndOr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": BoolOpStr_IsAndOr = True
End Select
End Function

Function BoolOpStr_IsEqNe(A$) As Boolean
Select Case UCase(A)
Case "EQ", "NE": BoolOpStr_IsEqNe = True
End Select
End Function

Function BoolOpStr_IsVdt(A$) As Boolean
BoolOpStr_IsVdt = IsInUCaseSy(A, BoolOpSy)
End Function

Function Bool_Txt_IfTrue$(A As Boolean, T$)
If A Then Bool_Txt_IfTrue = T
End Function

Property Get BoolOpSy() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = Ssl_Sy("AND OR")
End If
BoolOpSy = Y
End Property
