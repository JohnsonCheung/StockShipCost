Attribute VB_Name = "MIde_Mth_IsSel"
Option Compare Binary
Option Explicit

Function Mth3Nm_IsSel(MthNm$, ShtTy$, ShtMdy$, A As WhMth) As Boolean
If IsNothing(A) Then Mth3Nm_IsSel = True: Exit Function
If Not MthShtMdy_IsSel(ShtMdy, A.InMdy) Then Exit Function
If Not Itm_IsSel_ByAy(MthShtTy_MthShtKd(ShtTy), A.InShtKd) Then Exit Function
Mth3Nm_IsSel = Nm_IsSel(MthNm, A.Nm)
End Function

Function MthNmBrk_IsSel(MthNmBrk$(), B As WhMth) As Boolean
Select Case Sz(MthNmBrk)
Case 0: Exit Function
Case 3: MthNmBrk_IsSel = Mth3Nm_IsSel(MthNmBrk(0), MthNmBrk(1), MthNmBrk(2), B)
Case Else: Stop
End Select
End Function

Function MthShtKd_IsSel(Kd$, WhKd$) As Boolean
MthShtKd_IsSel = Itm_IsSel_ByAy(Kd, CvWhKd(WhKd))
End Function

Function MthShtMdy_IsSel(A$, ShtMdyAy$()) As Boolean
If Sz(ShtMdyAy) = 0 Then MthShtMdy_IsSel = True: Exit Function
Dim ShtMdy
For Each ShtMdy In ShtMdyAy
    If ShtMdy = "Pub" Then
        If A = "" Then MthShtMdy_IsSel = True: Exit Function
    End If
    If A = ShtMdy Then MthShtMdy_IsSel = True: Exit Function
Next
End Function

Private Sub ZZ()
Dim A$
Dim B As WhMth
Dim C$()
Mth3Nm_IsSel A, A, A, B
MthNmBrk_IsSel C, B
MthShtKd_IsSel A, A
MthShtMdy_IsSel A, C
End Sub

Private Sub Z()
End Sub
