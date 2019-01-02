Attribute VB_Name = "MIde_Fxa"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde__Fxa."
Property Get TmpFxaPj() As VBProject
Dim A As excel.Application
Set A = New_Xls
Set TmpFxaPj = XlsFxa_Pj(A, TmpFxa)
End Property
Function TmpFxa$(Optional Fdr$, Optional Fnn$)
TmpFxa = TmpFfn(".xlam", Fdr, Fnn)
End Function

Function Fxa_XCrt(Fxa) As VBProject
Set Fxa_XCrt = XlsFxa_XCrt(CurXls, Fxa)
End Function

Function XlsFxa_XCrt(A As excel.Application, Fxa) As VBProject
Ffn_XDltIfExist Fxa
Wb_XSavAs(A.Workbooks.Add, Fxa, XlFileFormat.xlOpenXMLAddIn).Close False
Dim Wb As Workbook
Set Wb = A.Workbooks.Open(Fxa)
Dim O As VBProject
Set O = VbePj_FfnPj(A.Vbe, Fxa)
O.Name = Ffn_Fnn(Fxa)
Wb.Save
Set XlsFxa_XCrt = O
End Function

Function XlsFxa_Pj(A As excel.Application, Fxa) As VBProject
Const CSub$ = CMod & "FxaPj"
If Not IsFxa(Fxa) Then XThw CSub, "Given [Fxa] is not ends with .xlam", A
Set XlsFxa_Pj = VbePj_Ffn(A.Vbe, Fxa)
End Function

Function FxaPj_Nm$(A)
FxaPj_Nm = Ffn_Fnn(A)
End Function

Function IsFxa(A) As Boolean
IsFxa = LCase(Ffn_Ext(A)) = ".xlam"
End Function

Function Pj_IsFxa(A As VBProject) As Boolean
Pj_IsFxa = IsFxa(Pj_Ffn(A))
End Function

Sub SrcPth_XBld_Fxa(SrcPth$)
Dim P As VBProject
   Dim Fnn$, F$
   Fnn = Ffn_Fnn(XRmv_LasChr(SrcPth))
   F = SrcPth & Fnn & ".xlam"
   Set P = XlsFxa_Pj(CurXls, F)
Dim SrcFfnAy$()
   Dim S
   SrcFfnAy = Ay_XWh_LikAy(Pth_FfnAy(SrcPth), Ssl_Sy("*.bas *.cls"))
   For Each S In SrcFfnAy
       P.ImpSrcFfn S
   Next
Pj_XRmv_OptCmpDbLin P
Pj_XSet_Rf_ByCfgFil P, SrcPth
Pj_XSav P
End Sub


Private Sub Z()
MIde__Fxa:
End Sub
