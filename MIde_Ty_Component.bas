Attribute VB_Name = "MIde_Ty_Component"
Option Compare Binary
Option Explicit
Function CvCmpTyAy(CmpTyAy0$) As vbext_ComponentType()
Dim X, O() As vbext_ComponentType
For Each X In Ssl_Sy(CmpTyAy0)
    Push O, CmpTy_ShtNmTy(X)
Next
CvCmpTyAy = O
End Function


Function CmpTy_ShtNmTy(Sht) As vbext_ComponentType
Select Case Sht
Case "Doc": CmpTy_ShtNmTy = vbext_ComponentType.vbext_ct_Document
Case "Cls": CmpTy_ShtNmTy = vbext_ComponentType.vbext_ct_ClassModule
Case "Std": CmpTy_ShtNmTy = vbext_ComponentType.vbext_ct_StdModule
Case "Frm": CmpTy_ShtNmTy = vbext_ComponentType.vbext_ct_MSForm
Case "ActX": CmpTy_ShtNmTy = vbext_ComponentType.vbext_ct_ActiveXDesigner
Case Else: Stop
End Select
End Function

Function CmpTy_ShtNm$(A As vbext_ComponentType)
Select Case A
Case vbext_ComponentType.vbext_ct_Document:    CmpTy_ShtNm = "Doc"
Case vbext_ComponentType.vbext_ct_ClassModule: CmpTy_ShtNm = "Cls"
Case vbext_ComponentType.vbext_ct_StdModule:   CmpTy_ShtNm = "Std"
Case vbext_ComponentType.vbext_ct_MSForm:      CmpTy_ShtNm = "Frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: CmpTy_ShtNm = "ActX"
Case Else: Stop
End Select
End Function
