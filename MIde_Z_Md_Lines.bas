Attribute VB_Name = "MIde_Z_Md_Lines"
Option Compare Binary
Option Explicit
Function Md_Lines$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Function
Md_Lines = A.Lines(1, A.CountOfLines)
End Function

Function MdFmCnt_Lines$(A As CodeModule, FmCnt As FmCnt)
With FmCnt
    If .Cnt <= 0 Then Exit Function
    MdFmCnt_Lines = A.Lines(.FmLno, .Cnt)
End With
End Function
Function Md_Ly(A As CodeModule) As String()
Md_Ly = SplitCrLf(Md_Lines(A))
End Function

Function MdFTNo_Lines$(A As CodeModule, B As FTNo)
Dim Cnt%: Cnt = FTNo_Cnt(B)
If Cnt = 0 Then Exit Function
MdFTNo_Lines = A.Lines(B.FmNo, Cnt)
End Function

Function MdFTNo_Ly(A As CodeModule, B As FTNo) As String()
MdFTNo_Ly = SplitCrLf(MdFTNo_Lines(A, B))
End Function

Function MdRe_Ly(A As CodeModule, B As RegExp) As String()
Dim Ix&(): Ix = AyRe_IxAy(Md_Ly(A), B)
Dim O$(), I, Md As CodeModule
Dim N$: N = Md_Nm(A)
If Sz(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, QQ_Fmt("Md_XGoLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
MdRe_Ly = O
End Function
