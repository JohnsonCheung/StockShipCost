Attribute VB_Name = "MIde_Z_Md_Lines_Bdy"
Option Compare Binary
Option Explicit
Function Md_BdyLines$(A As CodeModule)
If Md_XHas_NoMth(A) Then Exit Function
Md_BdyLines = A.Lines(Md_BdyFmLno(A), A.CountOfLines)
End Function

Function Md_BdyFmLno%(A As CodeModule)
Md_BdyFmLno = Md_DclLinCnt(A) + 1
End Function

Function MdBdyFmCnt(A As CodeModule) As FmCnt
Dim Lno&
Dim Cnt&
Lno = Md_BdyFmLno(A)
MdBdyFmCnt = New_FmCnt(Lno, Cnt)
End Function

Function Md_BdyLy(A As CodeModule) As String()
Md_BdyLy = SplitCrLf(Md_BdyLines(A))
End Function
