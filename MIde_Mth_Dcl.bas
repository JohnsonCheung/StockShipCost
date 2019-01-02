Attribute VB_Name = "MIde_Mth_Dcl"
Option Compare Binary
Option Explicit

Property Get CurMd_MthLinAy() As String()
CurMd_MthLinAy = Md_MthLinAy(CurMd)
End Property

Property Get CurPj_MthLinAy() As String()
CurPj_MthLinAy = Pj_MthLinAy(CurPj)
End Property

Function MdMthNm_MthLin$(A As CodeModule, MthNm$)
MdMthNm_MthLin = SrcMthNm_MthLin(Md_Src(A), MthNm)
End Function

Function Md_MthLinAy_Pub(A As CodeModule) As String()
Md_MthLinAy_Pub = Src_MthLinAy(Md_Src(A), WhMth_Pub)
End Function

Function Md_MthLinAy(A As CodeModule, Optional B As WhMth) As String()
Md_MthLinAy = Src_MthLinAy(Md_Src(A), B)
End Function

Function Mth_MthLin$(A As Mth)
Mth_MthLin = SrcMthNm_MthLin(Md_BdyLy(A.Md), A.Nm)
End Function

Function Pj_MthLinAy(A As VBProject) As String()
Dim I
For Each I In Pj_MdAy(A)
    PushIAy Pj_MthLinAy, Md_MthLinAy(CvMd(I))
Next
End Function

Function SrcMthNm_MthLin$(A$(), MthNm$)
SrcMthNm_MthLin = SrcIx_ContLin(A, SrcMthNm_MthIx(A, MthNm))
End Function

Function Src_MthLinAy(A$(), Optional B As WhMth) As String()
Dim J&
If IsNothing(B) Then
    For J = 0 To UB(A)
        If Lin_IsSel_WhMth(A(J), B) Then
            Push Src_MthLinAy, SrcIx_ContLin(A, J)
        End If
    Next
Else
    For J = 0 To UB(A)
        If Lin_IsMth(A(J)) Then
            Push Src_MthLinAy, SrcIx_ContLin(A, J)
        End If
    Next
End If
End Function

Private Sub Z_Src_MthLinDry()
Dry_XBrw Src_MthLinDry(Md_Src(CurMd))
End Sub

Function Src_MthLinDry(A$(), Optional B As WhMth) As Variant()
Dim I&, L$, O()
For I = 0 To UB(A)
    L = SrcIx_ContLin(A, I)
    O = Lin_MthLinDr(L)
    If Sz(O) > 0 Then
        PushI O, I + 1
        PushI O, SrcMthIx_MthNLin(A, I)
        PushI Src_MthLinDry, O
    End If
Next
End Function

Function Src_MthLinAy_Pub(A$()) As String()
Src_MthLinAy_Pub = Src_MthLinAy(A, WhMth_Pub)
End Function

Private Sub Z_Src_MthLinAy_Pub()
Dim MthNy$(), Src$()
Src = CurSrc
MthNy = Ap_Sy("Src_MthLinDry", "Mth_MthLin")
Ept = Ap_Sy("Function Mth_MthLin$(A As Mth)", "Function Src_MthLinDry(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = Src_MthLinAy_Pub(Src)
    C
    Return
End Sub
