Attribute VB_Name = "MIde_Mth_Lin_Is"
Option Compare Binary
Option Explicit
Function Lin_IsPrp(A) As Boolean
Lin_IsPrp = Lin_MthKd(A) = "Property"
End Function
Private Sub Z_Lin_IsMth()
GoTo ZZ
Dim A$
A = "Function Lin_IsMth(A, Optional B As WhMth) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = Lin_IsMth(A)
    C
    Return
ZZ:
Dim L, O$()
For Each L In CurSrc
    If Lin_IsMth(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub

Function Lin_IsMth(A) As Boolean
Lin_IsMth = Lin_MthKd(A) <> ""
End Function

Function Lin_IsSel_WhMth(A, B As WhMth) As Boolean
Lin_IsSel_WhMth = MthNmBrk_IsSel(Lin_MthNmBrk(A), B)
End Function
Private Sub Z_Lin_IsPubMth()
Dim O$(), L
For Each L In Pj_Src(CurPj)
    If Lin_IsPubMth(L) Then PushI O, L
Next
Brw O
End Sub
Function Lin_IsPubMth(A) As Boolean
Dim L$: L = A
If Not Ay_XHas(Array("", "Public"), XShf_MthMdy(L)) Then Exit Function
Lin_IsPubMth = XShf_MthShtTy(L) <> ""
End Function

Private Sub Z()
Z_Lin_IsMth
MIde_MthLin_Is:
End Sub
