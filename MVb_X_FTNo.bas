Attribute VB_Name = "MVb_X_FTNo"
Option Compare Binary
Option Explicit
Function FTNoAy_Ly(A() As FTNo) As String()
Dim I
For Each I In AyNz(A)
    PushI FTNoAy_Ly, FTNo_Str(CvFTNo(I))
Next
End Function
Sub FTNo_XDmp(A As FTNo)
D FTNo_Str(A)
End Sub
Function FTNo_Str$(A As FTNo)
FTNo_Str = QQ_Fmt("FTNo(? ?)", A.FmNo, A.ToNo)
End Function

Function FTNoAy_Cnt&(A() As FTNo)
Dim O&, M
For Each M In A
    O = O + FTNo_Cnt(CvFTNo(M))
Next
End Function

Function FTNo_FTIx(A As FTNo) As FTIx
Set FTNo_FTIx = New_FTIx(A.FmNo - 1, A.ToNo - 1)
End Function

Function FTIx_FTNo(A As FTIx) As FTNo
Set FTIx_FTNo = New_FTNo(A.FmIx + 1, A.ToIx + 1)
End Function

Function FTNo_Cnt&(A As FTNo)
Dim O&
O = A.ToNo - A.FmNo + 1
If O < 0 Then Stop
FTNo_Cnt = O
End Function

Function New_FTNo(FmNo&, ToNo&) As FTNo
Dim O As New FTNo
Set New_FTNo = O.Init(FmNo, ToNo)
End Function

Function CvFTNo(A) As FTNo
Set CvFTNo = A
End Function

