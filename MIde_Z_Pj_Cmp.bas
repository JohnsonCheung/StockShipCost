Attribute VB_Name = "MIde_Z_Pj_Cmp"
Option Compare Binary
Option Explicit
Function PjClsAndModNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
PjClsAndModNy = Pj_CmpNy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function

Function PjClsAndModAy(A As VBProject, Optional Patn$, Optional Exl$) As CodeModule()
PjClsAndModAy = PjModAy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function

Function PjClsAndModCmpAy(A As VBProject, Optional Patn$, Optional Exl$) As VBComponent()
PjClsAndModCmpAy = PjCmpAy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function
Function Pj_ClsAndModNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
Pj_ClsAndModNy = Itr_Ny(PjClsAndModCmpAy(A, Patn, Exl))
End Function

Function PjClsAy(A As VBProject, Optional B As WhNm) As CodeModule()
PjClsAy = Pj_MdAy(A, WhMd("Cls"))
End Function

Function Pj_ClsNy(A As VBProject, Optional B As WhNm) As String()
Pj_ClsNy = Pj_CmpNy(A, WhMd("Cls", B))
End Function

Function PjCmp(A As VBProject, Nm) As VBComponent
Set PjCmp = A.VBComponents(Nm)
End Function

Function PjCmpAy(A As VBProject, Optional B As WhMd) As VBComponent()
Dim I
For Each I In AyNz(Pj_MdAy(A, B))
    PushObj PjCmpAy, CvMd(I).Parent
Next
End Function

Function Pj_CmpNy(A As VBProject, Optional B As WhMd) As String()
Pj_CmpNy = Itr_Ny(PjCmpAy(A, B))
End Function

Function PjFstMbr(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    Set PjFstMbr = Cmp.CodeModule
    Exit Function
Next
End Function

Function PjFstMd(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_StdModule Then
        Set PjFstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function Pj_XHas_CmpNm(A As VBProject, Nm) As Boolean
Pj_XHas_CmpNm = Itr_XHas_Nm(A.VBComponents, Nm)
End Function

Function Pj_XHas_CmpNm_WhRe(A As VBProject, Re As RegExp) As Boolean
Pj_XHas_CmpNm_WhRe = Itr_XHas_Nm_WhRe(A.VBComponents, Re)
End Function
Function Pj_XHas_MdNm(A As VBProject, Nm) As Boolean
Pj_XHas_MdNm = Pj_XHas(A, Nm, vbext_ct_StdModule)
End Function

Private Function Pj_XHas(A As VBProject, Nm, Ty As vbext_ComponentType) As Boolean
Dim T As vbext_ComponentType
If Not Itr_XHas_Nm(A.VBComponents, Nm) Then Exit Function
T = PjCmp(A, Nm).Type
If T = Ty Then Pj_XHas = True: Exit Function
XDmp_Ly CSub, "Pj has Cmp not as expected type", "PjCmpNm EptTy ActTy", Pj_Nm(A), Nm, CmpTy_ShtNm(Ty), CmpTy_ShtNm(T)
End Function

Function Pj_XHasCls(A As VBProject, Nm) As Boolean
Pj_XHasCls = Pj_XHas(A, Nm, vbext_ct_ClassModule)
End Function

Function Pj_XHasNoStdClsMd(A As VBProject) As Boolean
Dim C As VBComponent
For Each C In A.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
Pj_XHasNoStdClsMd = True
End Function

Function Pj_MdAy(A As VBProject, Optional B As WhMd) As CodeModule()
If IsNothing(B) Then
    Pj_MdAy = ItrPrp_Into(A.VBComponents, "CodeModule", Pj_MdAy)
    Exit Function
End If
Dim C
For Each C In AyNz(Itr_XWh_Nm(A.VBComponents, B.Nm))
    With CvCmp(C)
        If Itm_IsSel_ByAy(.Type, B.InCmpTy) Then
            PushObj Pj_MdAy, .CodeModule
        End If
    End With
Next
End Function

Function PjModAy_PATN(A As VBProject, Patn$) As CodeModule()
PjModAy_PATN = PjModAy(A, WhNm(Patn))
End Function

Function PjModAy_PFX(A As VBProject, Pfx) As CodeModule()
PjModAy_PFX = PjModAy(A, WhNm("^" & Pfx))
End Function

Function PjModAy(A As VBProject, Optional B As WhNm) As CodeModule()
PjModAy = Pj_MdAy(A, WhMd("Std", B))
End Function

Function PjModClsNy(A As VBProject, Optional B As WhNm) As String()
PjModClsNy = Pj_CmpNy(A, WhMd("Std Cls", B))
End Function

Function PjModNy(A As VBProject, Optional B As WhNm) As String()
PjModNy = Pj_CmpNy(A, WhMd("Std", B))
End Function


Private Sub Z_Pj_ClsNy()
Ay_XDmp Pj_ClsNy(CurPj)
End Sub

Function Pj_MdNy(A As VBProject, Optional B As WhMd) As String()
Pj_MdNy = Pj_CmpNy(A, B)
End Function

Function Pj_MthDDNy(A As VBProject, Optional B As WhMdMth) As String()
Dim WhMth As WhMth
    Set WhMth = WhMdMth_WhMth(B)
Dim I
For Each I In Pj_MdAy(A, WhMdMth_WhMd(B))
    PushIAy Pj_MthDDNy, MthDDNy_XWh(Md_MthDDNy(CvMd(I), AddMdNm:=True), WhMth)
Next
End Function

Private Sub Z_Pj_MdAy()
Dim O() As CodeModule
O = Pj_MdAy(CurPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print Md_Nm(Md)
Next
End Sub

Private Sub Z_Pj_MdNy()
Ay_XDmp Pj_MdNy(CurPj)
End Sub


Private Sub Z()
Z_Pj_ClsNy
Z_Pj_MdAy
Z_Pj_MdNy
MIde_Z_Pj_Cmp:
End Sub
