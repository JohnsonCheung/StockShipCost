Attribute VB_Name = "MIde_Mth_Nm_Brk"
Option Compare Binary
Option Explicit

Sub Mth_MthNmBrk_XAsg(A As Mth, OMdy$, OMthTy$)
Dim L$
L = Mth_MthLin(A)
OMdy = TakMdy(L)
OMthTy = LinMthTy(L)
End Sub

Function MthNmBrkAy_MthDDNy(A() As Variant) As String()
MthNmBrkAy_MthDDNy = DryJnDotSy(A)
End Function

Function MthNmBrkAy_XWh_Dup(A()) As Variant()
'MthBrk is Sy of Mdy Ty Nm
Dim Dry(): Dry = Dry_XWh_ColInAy(A, 0, Array("", "Public")) '
MthNmBrkAy_XWh_Dup = Dry_XWh_ColHasDup(Dry, 2)
End Function

Function MthNmBrk_Nm$(MthNmBrk$())
Select Case Sz(MthNmBrk)
Case 0:
Case 3: MthNmBrk_Nm = MthNmBrk(0)
Case Else: Stop
End Select
End Function

Function MthNmBrkAy_XWh(A() As Variant, B As WhMth) As Variant()
Dim Brk
For Each Brk In AyNz(A)
    If MthNmBrk_IsSel(CvSy(Brk), B) Then PushI MthNmBrkAy_XWh, Brk
Next
End Function

Function Lin_MthNmBrk(A, Optional B As WhMth) As String()
'Return Ay of (Nm,Ty,Kd)
Dim O$()
O = XShf_MthNmBrk(CStr(A))
Lin_MthNmBrk = XShf_MthNmBrk(CStr(A))
If MthNmBrk_IsSel(O, B) Then Lin_MthNmBrk = O
End Function

Function MthNmBrkAyNy(A() As Variant) As String()
MthNmBrkAyNy = DryDistSy(A, 2)
End Function

Sub Lin_MthNmBrkAsg(A$, OMdy$, OTy$, ONm$)
Dim L$: L = A
OMdy = XShf_MthMdy(L)
OTy = XShf_MthShtTy(L)
ONm = XTak_Nm(L)
End Sub
