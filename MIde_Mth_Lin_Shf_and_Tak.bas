Attribute VB_Name = "MIde_Mth_Lin_Shf_and_Tak"
Option Compare Binary
Option Explicit

Function XShf_ItmNy(A$, ItmNy0) As Variant()
XShf_ItmNy = Ay_XShf_ItmNy(Lin_TermAy(A), ItmNy0)
End Function
Function XShf_MthShtTy$(OLin)
Dim O$: O = XShf_MthTy(OLin): If O = "" Then Exit Function
XShf_MthShtTy = MthTy_MthShtTy(O)
End Function
Function XShf_MthTy$(OLin)
Dim O$: O = XTak_MthTy(OLin): If O = "" Then Exit Function
XShf_MthTy = O
OLin = LTrim(XRmv_Pfx(OLin, O))
End Function

Sub XShf_MthShtTy_Asg(A, OMthTy, ORst$)
AyAsg XShf_MthShtTy(A), OMthTy, ORst
End Sub

Function XShf_As(A) As Variant()
Dim L$
L = LTrim(A)
If Left(L, 3) = "As " Then XShf_As = Array(True, LTrim(Mid(L, 4))): Exit Function
XShf_As = Array(False, A)
End Function

Function XShf_MthShtMdy$(OLin)
Dim O$: O = XShf_MthMdy(OLin): If O = "" Then Exit Function
XShf_MthShtMdy = MthMdy_MthShtMdy(O)
End Function

Function XShf_MthMdy$(OLin)
Dim O$
O = TakMdy(OLin): If O = "" Then Exit Function
XShf_MthMdy = O
OLin = LTrim(XRmv_Pfx(OLin, O))
End Function

Function XShf_MthNmBrk(OLin) As String()
Dim B$()
ReDim B$(2)
B(2) = XShf_MthShtMdy(OLin)
B(1) = XShf_MthShtTy(OLin): If B(1) = "" Then Exit Function
B(0) = XShf_Nm(OLin)
XShf_MthNmBrk = B
End Function

Function XShf_Kd$(OLin)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
XShf_Kd = T
OLin = LTrim(XRmv_Pfx(OLin, T))
End Function

Function XShf_MthSfx$(OLin)
XShf_MthSfx = XShf_Chr(OLin, "#!@#$%^&")
End Function

Function XShf_Nm$(OLin)
Dim O$: O = XTak_Nm(OLin): If O = "" Then Exit Function
XShf_Nm = O
OLin = XRmv_Pfx(OLin, O)
End Function

Function XShf_Rmk(A) As String()
Dim L$
L = LTrim(A)
If XTak_FstChr(L) = "'" Then
    XShf_Rmk = Ap_Sy(Mid(L, 2), "")
Else
    XShf_Rmk = Ap_Sy("", A)
End If
End Function

Function TakMdy$(A)
TakMdy = XTak_PfxAySpc(A, MdyAy)
End Function

Function TakMthKd$(A)
TakMthKd = XTak_PfxAySpc(A, MthKdAy)
End Function

Function XTak_MthTy$(A)
XTak_MthTy = XTak_PfxAySpc(A, MthTyAy)
End Function

