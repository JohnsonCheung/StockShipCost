Attribute VB_Name = "MIde_Mth_Ix"
Option Compare Binary
Option Explicit
Private Sub Z_Src_MthIx()
Dim IxAy&(), Src$(), Ix
Src = Md_Src(Md("AAAMod"))
IxAy = Src_MthIxAy(Src)
For Each Ix In IxAy
    If Lin_MthKd(Src(Ix)) = "" Then
        Debug.Print Ix
        Debug.Print Src(Ix)
    End If
Next
End Sub

Function Mth_FmCnt(A As Mth) As FmCnt()
Mth_FmCnt = SrcMthNm_FmCntAy(Md_BdyLy(A.Md), A.Nm)
End Function

Function Src_MthIxAy(A$(), Optional B As WhMth) As Long()
Dim L, J&
If IsNothing(B) Then
    For Each L In AyNz(A)
        If Lin_IsMth(L) Then
            PushI Src_MthIxAy, J
        End If
        J = J + 1
    Next
Else
    For Each L In AyNz(A)
        If Lin_IsSel_WhMth(L, B) Then
            PushI Src_MthIxAy, J
        End If
        J = J + 1
    Next
End If
End Function

Function SrcMthNm_MthIxAy(A$(), MthNm) As Long()
Dim L, J&, Ix&
Ix = SrcMthNm_MthIx(A, MthNm): If Ix = -1 Then Exit Function
PushI SrcMthNm_MthIxAy, Ix
If Lin_IsPrp(A(Ix)) Then
    Ix = SrcMthNm_MthIx(A, MthNm, Ix + 1)
    If Ix > 0 Then
        PushI SrcMthNm_MthIxAy, Ix
    End If
End If
End Function

Function SrcMthNm_MthIx&(A$(), MthNm, Optional Fm& = 0)
Dim I
For I = Fm To UB(A)
    If Lin_MthNm(A(I)) = MthNm Then
        SrcMthNm_MthIx = I
        Exit Function
    End If
Next
SrcMthNm_MthIx = -1
End Function
Function SrcMthIx_MthIxTo&(A$(), MthIx)
Dim T$, F$, J&
T = Lin_MthKd(A(MthIx)): If T = "" Then Stop
F = "End " & T
If HasSubStr(A(MthIx), F) Then SrcMthIx_MthIxTo = MthIx: Exit Function
For J = MthIx + 1 To UB(A)
    If XHas_Pfx(LTrim(A(J)), F) Then SrcMthIx_MthIxTo = J: Exit Function
Next
Stop
End Function

Private Sub Z_Src_MthIx1()
Dim A$(), Ix&(), O$(), I
A = CurSrc
Ix = Src_MthIxAy(CurSrc)
For Each I In Ix
    PushI O, A(I)
Next
Brw O
End Sub

Function Src_FstMthIx&(A$())
Dim J&
For J = 0 To UB(A)
   If Lin_IsMth(A(J)) Then
       Src_FstMthIx = J
       Exit Function
   End If
Next
Src_FstMthIx = -1
End Function
Function MdMthNm_Lno&(A As CodeModule, MthNm)
MdMthNm_Lno = 1 + SrcMthNm_MthIx(Md_Src(A), MthNm)
End Function
Function MdMthNm_LnoAy(A As CodeModule, MthNm) As Long()
MdMthNm_LnoAy = Ay_XAdd_1(SrcMthNm_MthIxAy(Md_Src(A), MthNm))
End Function




Private Sub Z()
Z_Src_MthIx
Z_Src_MthIx1
MIde_Mth_Ix:
End Sub
