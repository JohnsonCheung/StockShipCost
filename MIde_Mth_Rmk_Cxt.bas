Attribute VB_Name = "MIde_Mth_Rmk_Cxt"
Option Compare Binary
Option Explicit

Private Function SrcMthFTIx_MthCxtFTIx(A$(), B As FTIx) As FTIx
Dim FmIx&
    Dim Ix&
    For Ix = B.FmIx To B.ToIx
        If Not XTak_LasChr(A(Ix)) = "_" Then
            FmIx = Ix + 1
            GoTo Fnd
        End If
    Next
    XThw CSub, "All lines with FTIx of Src is with [_] at end", "Src MthFTIx", A, FTIx_Str(B)
Fnd:
Dim ToIx&
    ToIx = B.ToIx - 1
Set SrcMthFTIx_MthCxtFTIx = New_FTIx(FmIx, ToIx)
End Function

Function Mth_MthCxtFTNoAy(A As Mth) As FTNo()
Mth_MthCxtFTNoAy = SrcMthNm_MthCxtFTNoAy(Md_Src(A.Md), A.Nm)
End Function

Sub MdFTNo_XRmk(A As CodeModule, B As FTNo)
If Src_IsAllRmked(MdFTNo_Ly(A, B)) Then Exit Sub
Dim J%, L$
For J = B.FmNo To B.ToNo
    L = A.Lines(J, 1)
    A.ReplaceLine J, "'" & L
Next
End Sub

Sub MdFTNo_XUmRmk(A As CodeModule, B As FTNo)
If Not Src_IsAllRmked(MdFTNo_Ly(A, B)) Then Exit Sub
Dim J%, L$
For J = B.FmNo To B.ToNo
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then XThw CSub, "Program Error: Src_IsAllRmked return when some line is not remarked", "The-Not-Remarked-line Md FTNo", L, Md_Nm(A), FTNo_Str(B)
    A.ReplaceLine J, Mid(L, 2)
Next
End Sub

Function Src_IsAllRmked(A$()) As Boolean
Dim L
For Each L In AyNz(A)
    If Left(L, 1) <> "'" Then Exit Function
Next
Src_IsAllRmked = True
End Function

Function SrcMthNm_MthCxtFTNoAy(A$(), MthNm$) As FTNo()
Dim I
For Each I In AyNz(SrcMthNm_FTIxAy(A, MthNm))
    Dim B As FTIx
    Dim C As FTNo
        Set B = SrcMthFTIx_MthCxtFTIx(A, CvFTIx(I))
        Set C = FTIx_FTNo(B)
    PushObj SrcMthNm_MthCxtFTNoAy, C
Next
End Function

Private Sub ZZ_SrcMthNm_CxtFTNo _
 _
()

Dim I
For Each I In Mth_MthCxtFTNoAy(CurMth)
    With CvFTNo(I)
        Debug.Print .FmNo, .ToNo
    End With
Next
End Sub

