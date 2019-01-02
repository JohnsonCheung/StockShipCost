Attribute VB_Name = "MVb_X_FTIx"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_X_FTIx."
Function FTIxAy_FmCntAy(A() As FTIx) As FmCnt()
Dim I
For Each I In AyNz(A)
    PushObj FTIxAy_FmCntAy, FTIx_FmCnt(CvFTIx(A))
Next
End Function

Function FTIx_IsEmp(A As FTIx) As Boolean
FTIx_IsEmp = True
If A.FmIx < 0 Then Exit Function
If A.ToIx < 0 Then Exit Function
If A.FmIx > A.ToIx Then Exit Function
FTIx_IsEmp = False
End Function


Function FTIx_Cnt&(A As FTIx)
FTIx_XAss_Vdt A
FTIx_Cnt = A.ToIx - A.FmIx + 1
End Function
Function FTIx_Str$(A As FTIx)

End Function
Sub FTIx_XAss_Vdt(A As FTIx)
If Not FTIx_IsVdt(A) Then
    XThw CSub, "Invalid FTIx", "FTIx", FTIx_Str(A)
End If
End Sub

Function FTIx_XHas_U(A As FTIx, U&) As Boolean
If U < 0 Then Stop
If FTIx_IsEmp(A) Then Exit Function
If A.FmIx > U Then Exit Function
If A.ToIx > U Then Exit Function
FTIx_XHas_U = True
End Function

Function FTIx_IsVdt(A As FTIx) As Boolean
FTIx_IsVdt = True
If A.FmIx < 0 Then Exit Function
If A.ToIx < 0 Then Exit Function
If A.FmIx > A.ToIx Then Exit Function
FTIx_IsVdt = False
End Function

Sub FmToIxU_XAss(FmIx, ToIx, U)
Const CSub$ = CMod & "FmToIxU_XAss"
If FmIx < 0 Then XThw CSub, "[FmIx] is negative, where [U] and [ToIx]", "FmIx U ToIx", FmIx, U, ToIx
If ToIx < 0 Then XThw CSub, "[ToIx] is negative, where [U] and [FmIx]", "ToIx U FmIx", ToIx, U, FmIx
End Sub


Function FTIx_FmCnt(A As FTIx) As FmCnt
Dim Lno&, Cnt&
   Cnt = FTIx_Cnt(A)
   Lno = A.FmIx + 1
Set FTIx_FmCnt = New_FmCnt(Lno, Cnt)
End Function

Function FTIxNo(A As FTIx) As FTNo
Set FTIxNo = New_FTNo(A.FmIx + 1, A.ToIx + 1)
End Function

Function New_FTIx(FmIx, ToIx) As FTIx
Dim O As New FTIx
Set New_FTIx = O.Init(FmIx, ToIx)
End Function
Function CvFTIx(A) As FTIx
Set CvFTIx = A
End Function


