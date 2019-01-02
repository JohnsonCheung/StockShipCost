Attribute VB_Name = "MIde_Export"
Option Compare Binary
Option Explicit

Sub Md_XExp(A As CodeModule)
Dim F$: F = Md_SrcFfn(A)
A.Parent.Export F
Debug.Print Md_Nm(A)
End Sub

Sub Pj_XExp(A As VBProject)
Debug.Print "Pj_XExp: " & Pj_Nm(A) & "-----------------------------"
Dim P$
    P = Pj_SrcPth(A)
    If P = "" Then
        Debug.Print QQ_Fmt("Pj_XExp: Pj(?) does not have FileName", A.Name)
        Exit Sub
    End If
Pth_XClr_Fil P 'Clr SrcPth ---
Ffn_XCpy_ToPth A.FileName, P, OvrWrt:=True
'Export Mod
    Dim I
    For Each I In AyNz(Pj_MdAy(A)) ' Only Cls & Mod will be exported
        Md_XExp CvMd(I)  'Exp each md --
    Next
Pj_XExp_Rf A
End Sub

Sub Pj_XExp_Rf(A As VBProject)
Ay_XWrt Pj_RfLy_FmPj(A), Pj_RfCfgFfn(A)
End Sub

Sub Pj_XExp_Src(A As VBProject)
Pj_XCpy_ToSrcPth A
Pth_XClr_Fil Pj_SrcPth(A)
Dim Md As CodeModule, I
For Each I In PjModAy(A)
    Set Md = I
    Md_XExp Md
Next
End Sub

Private Sub ZZ()
Dim A As CodeModule
Dim B As VBProject
Dim XX
Md_XExp A
Pj_XExp B
Pj_XExp_Rf B
Pj_XExp_Src B
End Sub

Private Sub Z()
End Sub
