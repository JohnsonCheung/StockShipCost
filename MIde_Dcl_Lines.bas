Attribute VB_Name = "MIde_Dcl_Lines"
Option Compare Binary
Option Explicit

Private Sub Z_Src_DclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = Src_SrtLy(B1)
Dim A1%: A1 = Src_DclLinCnt(B1)
Dim A2%: A2 = Src_DclLinCnt(Src_SrtLy(B1))
End Sub
Function Src_DclLinCnt%(A$())
Dim I&
    I = Src_FstMthIx(A)
    If I = -1 Then
        I = UB(A) + 1
    Else
        I = SrcMthIx_MthIxTopRmkFm(A, I)
    End If
Dim O&
    For I = I - 1 To 0 Step -1
         If IsCdLin(A(I)) Then O = I + 1: GoTo X
    Next
    O = 0
X:
Src_DclLinCnt = O
End Function

Private Sub Z_DclTyNm_TyLines()
Debug.Print DclTyNm_TyLines(Md_DclLy(CurMd), "AA")
End Sub
Function Src_DclLines$(A$())
Src_DclLines = JnCrLf(Src_DclLy(A))
End Function

Function Src_DclLy(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim N&
   N = Src_DclLinCnt(A)
If N = 0 Then Exit Function
Src_DclLy = Ay_FstNEle(A, N)
End Function

Function Md_DclLinCnt%(A As CodeModule)
If A.CountOfDeclarationLines = 0 Then Exit Function
Md_DclLinCnt = Src_DclLinCnt(Md_Src(A))
End Function

Function MdDclLines$(A As CodeModule)
Dim Cnt%
Cnt = Md_DclLinCnt(A)
If Cnt = 0 Then Exit Function
MdDclLines = A.Lines(1, Cnt)
End Function

Function Md_DclLy(A As CodeModule) As String()
Md_DclLy = SplitCrLf(MdDclLines(A))
End Function


Private Sub Z()
Z_DclTyNm_TyLines
Z_Src_DclLinCnt
MIde_Dcl_Lines:
End Sub
