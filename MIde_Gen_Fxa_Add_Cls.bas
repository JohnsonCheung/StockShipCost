Attribute VB_Name = "MIde_Gen_Fxa_Add_Cls"
Option Compare Binary
Option Explicit

Private Function WCmpAy(Fm As VBProject, ToPj As VBProject, FxaNm) As VBComponent()
Dim Src$() ' CmpNy
Dim Tar$()
Dim N
Src = PjNm_ClsNy_OWNING(Pj_Nm(ToPj))
Tar = Pj_ClsNy(ToPj)
For Each N In AyNz(AyMinus(Src, Tar))
    PushObj WCmpAy, Fm.VBComponents(N)
Next
End Function

Sub CurPj_XWrt_ClsToFxa()
Pj_XWrt_ClsToFxa CurPj
End Sub

Function PjFxaNm_FxaPjAy(A As VBProject, X As excel.Application) As VBProject()
Dim P
    P = Pj_FxaPth(A)

Dim N ' FxaNm
Dim Fxa
Dim ToPj As VBProject
For Each N In AyNz(Pj_FxaNy(A))
    Fxa = P & N & ".xlam"
    If Ffn_Exist(Fxa) Then
        PushObj PjFxaNm_FxaPjAy, XlsFxa_XOpn_Pj(X, Fxa)
    End If
Next
End Function

Sub Pj_XWrt_ClsToFxa(A As VBProject)
Dim X As excel.Application
    Set X = New_Xls
Dim P, ToPj As VBProject, FxaNm$
For Each P In PjFxaNm_FxaPjAy(A, X)
    Set ToPj = P
    FxaNm = Ffn_Fnn(Pj_Ffn(ToPj))
    CmpAy_XCpy WCmpAy(A, ToPj, FxaNm), ToPj
    Pj_XSav ToPj
Next
Xls_XQuit X
End Sub

Sub XWrt_ClsToFxa()
CurPj_XWrt_ClsToFxa
End Sub

Private Sub Z_XWrt_ClsToFxa()
XWrt_ClsToFxa
End Sub

Private Sub ZZ()
Dim B As VBProject
CurPj_XWrt_ClsToFxa
Pj_XWrt_ClsToFxa B
XWrt_ClsToFxa
End Sub

Private Sub Z()
Z_XWrt_ClsToFxa
End Sub
