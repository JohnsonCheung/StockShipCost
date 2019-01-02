Attribute VB_Name = "MIde_Export_Nm"
Option Compare Binary
Option Explicit

Sub CurPj_SrcPth_XBrw()
Pth_XBrw Pj_SrcPth(CurPj)
End Sub

Property Get CurPj_SrcPth$()
CurPj_SrcPth = Pj_SrcPth(CurPj)
End Property

Function Md_SrcExt$(A As CodeModule)
Dim O$
Select Case A.Parent.Type
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "Md_SrcExt: Unexpected Md_CmpTy.  Should be [Class or Module or Document]"
End Select
Md_SrcExt = O
End Function

Function Md_SrcFfn$(A As CodeModule)
Md_SrcFfn = Pj_SrcPth(Md_Pj(A)) & Md_Fn(A)
End Function
