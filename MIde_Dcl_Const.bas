Attribute VB_Name = "MIde_Dcl_Const"
Option Compare Binary
Option Explicit

Function XShf_XConst(O) As Boolean
XShf_XConst = XShf_X(O, "Const")
End Function

Function Md_XHas_Const(A As CodeModule, ConstNm$) As Boolean
Dim J%
For J = 1 To A.CountOfDeclarationLines
    If Lin_ConstNm(A.Lines(J, 1)) = ConstNm Then Md_XHas_Const = True: Exit Function
Next
End Function

Function Lin_ConstNm$(A)
Dim L$: L = XRmv_Mdy(A)
If XShf_XConst(L) Then Lin_ConstNm = XTak_Nm(L)
End Function

