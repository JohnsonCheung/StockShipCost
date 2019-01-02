Attribute VB_Name = "MIde_Gen_Fxa_Add_Rf"
Option Compare Binary
Type RfGUIDLin
    Lin As String
End Type
Function RfGUIDLin(A) As RfGUIDLin
RfGUIDLin.Lin = A
End Function
Sub CurPj_XSet_Rf_ByRfFfn_TO_FXA()
Pj_XSet_Rf_ByRfFfn_TO_FXA CurPj
End Sub

Function FxaNmAddRf(FxaNm)
Dim P As VBProject
Set P = FxaNm_Pj(FxaNm)
Pj_XSet_Rf_ByRfFfn_USR P
Pj_XSet_Rf_ByRfFfn_STD P
End Function
Sub Pj_XSet_Rf_ByRfFfn_STD(A As VBProject)
Dim L
Ay = RfNmGUIDLy(Pj_Nm(A))
For Each L In AyNz(RfNmGUIDLy(Pj_Nm(A)))
    Pj_XSet_Rf_ByRfFfn_GUID_LIN A, L
Next
End Sub
Function StdGUIDLinUB%(A() As RfGUIDLin)
StdGUIDLinUB = StdGUIDLinSz(A) - 1
End Function
Function StdGUIDLinSz%(A() As RfGUIDLin)
On Error Resume Next
StdGUIDLinSz = UBound(A) + 1
End Function
Function FxaNm_RfNy_USR(FxaNm) As String()
FxaNm_RfNy_USR = Ssl_Sy(PjRfDfn_USR(FxaNm))
End Function
Sub Pj_XSet_Rf_ByRfFfn_USR(A As VBProject)
Dim RfNy$()
    RfNy = FxaNm_RfNy_USR(Pj_Nm(A))
Dim I, P$
P = Pj_Pth(A)
'FunMsgNyAp_XDmp CSub, "FxaNm's Fxa should exist", "[Usr Defined RfAy]", RfNy
For Each I In AyNz(RfNy)
    Pj_XSet_Rf_ByRfFfn A, I, P & I & ".xlam"
Next
End Sub

Sub Pj_XSet_Rf_ByRfFfn_TO_FXA(A As VBProject)
Dim P, Pj As VBProject
Dim X As excel.Application
Set X = New_Xls
For Each P In AyNz(PjFxaNm_FxaPjAy(A, X))
    Set Pj = P
    Pj_XSet_Rf_ByRfFfn_USR Pj
    Pj_XSet_Rf_ByRfFfn_STD Pj
Next
Xls_XQuit X
End Sub

Sub XAdd_Rf_TO_FXA()
CurPj_XSet_Rf_ByRfFfn_TO_FXA
End Sub

