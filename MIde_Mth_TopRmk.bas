Attribute VB_Name = "MIde_Mth_TopRmk"
Option Compare Binary
Option Explicit
Private Sub Z_MdMthNm_FmCntAyWithTopRmk()
Dim A As CodeModule, M, Ept() As FmCnt, Act() As FmCnt

Set A = Md("IdeMthFmCnt")
M = "Z_MdMthNm_FmCntAyWithTopRmk "
PushObj Ept, New_FmCnt(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MdMdMthNm_FmCntAyWithTopRmk(A, M)
    If Not FmCntAy_IsEq(Act, Ept) Then Stop
    Return
End Sub

Function MdMdMthNm_FmCntAyWithTopRmk(A As CodeModule, M) As FmCnt()
MdMdMthNm_FmCntAyWithTopRmk = SrcMthNm_FmCntAyWithTopRmk(Md_Src(A), M)
End Function

Function SrcMthNm_FmCntAyWithTopRmk(A$(), MthNm) As FmCnt()
Dim FmIx&, ToIx&, IFm, Fm&
For Each IFm In AyNz(SrcMthNm_MthIxAy(A, MthNm))
    Fm = IFm
    FmIx = SrcMthIx_MthIxTopRmkFm(A, Fm)
    ToIx = SrcMthIx_MthIxTo(A, Fm)
    PushObj SrcMthNm_FmCntAyWithTopRmk, New_FmCnt(FmIx + 1, ToIx - FmIx + 1)
Next
End Function
Function SrcMthIx_TopRmk$(A$(), MthIx&)
Dim O$(), J&, L$
Dim Fm&: Fm = SrcMthIx_MthIxTopRmkFm(A, MthIx)
For J = Fm To MthIx - 1
    L = A(J)
    If XTak_FstChr(L) = "'" Then
        If L <> "'" Then
            PushI O, L
        End If
    End If
Next
SrcMthIx_TopRmk = Join(O, vbCrLf)
End Function


Function SrcMthIx_MthIxTopRmkFm&(A$(), MthIx&)
Dim M1&
    Dim J&
    For J = MthIx - 1 To 0 Step -1
        If IsCdLin(A(J)) Then
            M1 = J
            GoTo M1IsFnd
        End If
    Next
    M1 = -1
M1IsFnd:
Dim M2&
    For J = M1 + 1 To MthIx - 1
        If Trim(A(J)) <> "" Then
            M2 = J
            GoTo M2IsFnd
        End If
    Next
    M2 = MthIx
M2IsFnd:
SrcMthIx_MthIxTopRmkFm = M2
End Function



Private Sub Z()
Z_MdMthNm_FmCntAyWithTopRmk
MIde_Mth_TopRmk:
End Sub
