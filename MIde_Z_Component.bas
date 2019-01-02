Attribute VB_Name = "MIde_Z_Component"
Option Compare Binary
Option Explicit

Function Cmp_IsCls(A As VBComponent) As Boolean
Cmp_IsCls = A.Type = vbext_ct_ClassModule
End Function

Function Cmp_IsClsOrStd(A As VBComponent) As Boolean
Select Case A.Type
Case vbext_ct_ClassModule, vbext_ct_StdModule: Cmp_IsClsOrStd = True
End Select
End Function
Function Cmp_Pj(A As VBComponent) As VBProject
Set Cmp_Pj = A.Collection.Parent
End Function
Function Cmp_PjNm$(A As VBComponent)
Cmp_PjNm = Cmp_Pj(A).Name
End Function

Sub Cmp_XRmv(A As VBComponent)
A.Collection.Remove A
End Sub

Property Get CurCmp() As VBComponent
Set CurCmp = CurMd.Parent
End Property

Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function
Function Cmp_XSet_Nm(A As VBComponent, Nm, Optional Fun$ = "Cmp_XSet_Nm") As VBComponent
Dim Pj As VBProject
Set Pj = Cmp_Pj(A)
If Pj_XHas_CmpNm(Pj, Nm) Then
    XThw Fun, "Cmp already exist", "Cmp Exist-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    XThw Fun, "CmpNm same as Pj_Nm", "CmpNm", Nm
End If
A.Name = Nm
Set Cmp_XSet_Nm = A
End Function
