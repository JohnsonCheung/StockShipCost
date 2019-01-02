Attribute VB_Name = "MIde__PrpFun"
Option Explicit
Dim Info$()
Function Lin_IsPrpFun(A) As Boolean
Dim L$, B$
L = XRmv_Mdy(A)
If XShf_MthShtTy(L) <> "Function" Then Exit Function
If XShf_Nm(L) = "" Then Exit Function
XShf_MthShtTyChr L
Lin_IsPrpFun = Left(L, 2) = "()"
End Function

Function Md_LnoAy_PrpFun(A As CodeModule) As Long()
Dim J&, L
For Each L In Md_Src(A)
    J = J + 1
    If Lin_IsPrpFun(L) Then PushI Md_LnoAy_PrpFun, J
Next
End Function

Function Md_Ly_PrpFun(A As CodeModule) As String()
Dim L
For Each L In AyNz(Md_LnoAy_PrpFun(A))
    PushI Md_Ly_PrpFun, A.Lines(L, 1)
Next
End Function

Sub Md_XEns_PrpFun(A As CodeModule, Optional WhatIf As Boolean)
Dim L
For Each L In AyNz(Md_LnoAy_PrpFun(A))
    XUpd A, L, WhatIf
Next
End Sub

Function Pj_Ly_PrpFun(A As VBProject) As String()
Dim I, M As CodeModule, Pfx$
For Each I In AyNz(Pj_MdAy(A))
    Set M = I
    Pfx = Md_Nm(M) & "."
    PushIAy Pj_Ly_PrpFun, Ay_XAdd_Pfx(Md_Ly_PrpFun(M), Pfx)
Next
End Function

Sub Pj_XEns_PrpFun(A As VBProject, Optional WhatIf As Boolean)
Dim I
Erase Info
For Each I In AyNz(Pj_MdAy(A))
    Md_XEns_PrpFun CvMd(I), WhatIf
Next
Brw Info
End Sub

Sub XEns_PrpFun()
Md_XEns_PrpFun CurMd
End Sub

Private Sub XUpd(A As CodeModule, Lno, Optional WhatIf As Boolean)
Dim OldLin$
Dim NewLin$
    OldLin = A.Lines(Lno, 1)
    NewLin = Replace(A.Lines(Lno, 1), "Function", "Property Get")
If Not WhatIf Then A.ReplaceLine Lno, NewLin
PushI Info, "XEns_PrpFun:XUpd NewLin: " & OldLin
PushI Info, "                 OldLin: " & NewLin
End Sub

Private Sub Z_Pj_Ly_PrpFun()
Brw Pj_Ly_PrpFun(CurPj)
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As CodeModule
Dim C As VBProject
Lin_IsPrpFun A
Md_LnoAy_PrpFun B
Md_Ly_PrpFun B
Md_XEns_PrpFun B
Pj_Ly_PrpFun C
Pj_XEns_PrpFun C
XEns_PrpFun
Z_Pj_Ly_PrpFun
End Sub

Private Sub Z()
End Sub
