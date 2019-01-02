Attribute VB_Name = "MTp_X_Gp"
Option Compare Binary
Option Explicit
Function Ly_Gp(Ly$()) As Gp
Set Ly_Gp = New_Gp(Ly_LnxAy(Ly))
End Function
Function New_Gp(A() As Lnx) As Gp
Set New_Gp = New Gp
With New_Gp
    .LnxAy = A
End With
End Function

Function CvGp(A) As Gp
Set CvGp = A
End Function

Function Gp_Ly(A As Gp) As String()
Gp_Ly = LnxAy_Ly(A.LnxAy)
End Function

Function Gp_XRmv_Rmk(A As Gp) As Gp
Dim B() As Lnx: B = A.LnxAy
Dim M As Lnx
Dim J&, O() As Lnx
For J = 0 To UB(B)
    M = B(J)
    If Not Lin_IsTpRmkLin(M.Lin) Then
        PushObj O, M
    End If
Next
Set Gp_XRmv_Rmk = New_Gp(O)
End Function
