Attribute VB_Name = "MIde__Identifier"
Option Compare Binary
Option Explicit
Function Lin_IdfNy(A) As String()
Dim PmNy$(): ' PmNy = Lin_PmNy(SrcIx_ContLin(A, 0))
Dim DimNy$(): ' DimNy = MthLyDimNy(A)
'Lin_ExtNy = AyMinusAp(StrIdentifierAy(JnSpc(A)), DimNy, PmNy)
End Function

Function Src_DimNy(A$()) As String()
Dim S
For Each S In AyNz(MthLyDimStmtAy(A))
'    PushIAy MthLyDimNy, DimStmtNy(S)
Next
End Function

Function Dim_VarNy(A) As String()
'AA 1: Dim A

End Function
Function MthLyDimStmtAy(A$()) As String()

End Function
Function Ly_IdfNy(A$()) As String()

End Function

Function StrNy1(A) As String()
Dim O$: O = RplPun(A)
Dim O1$(): O1 = Ay_XWh_SingleEle(Ssl_Sy(O))
Dim O2$()
Dim J%
For J = 0 To UB(O1)
    If Not IsDigit(XTak_FstChr(O1(J))) Then Push O2, O1(J)
Next
StrNy1 = O2
End Function

Function StrNy(A) As String()
Dim O$, J%
O = A
Const C$ = "~!`@#$%^&*()-_=+[]{};;""'<>,.?/" & vbCr & vbLf
For J = 1 To Len(C)
    O = Replace(O, Mid(C, J, 1), " ")
Next
StrNy = Ay_XWh_Dist(Ssl_Sy(O))
End Function
Property Get VbKwAy() As String()
Static X$()
If Sz(X) = 0 Then
    X = Ssl_Sy("Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text")
End If
VbKwAy = X
End Property
