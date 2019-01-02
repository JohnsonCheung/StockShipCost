Attribute VB_Name = "MIde_Z_Pj_Cmp_Add"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Z_Pj_Cmp_Add."

Sub XAdd_Cls(Nm$)
Pj_XAdd_Cmp CurPj, Nm, vbext_ComponentType.vbext_ct_ClassModule
End Sub

Sub XAdd_Fun(FunNm$)
'Des: Add Empty-Fun-Mth to CurMd
Md_XApp_Lines CurMd, QQ_Fmt("Function ?()|End Function", FunNm)
MdMth_XGo CurMd, FunNm
End Sub

Sub XAdd_Mod(Nm$)
Pj_XAdd_Md CurPj, Nm
End Sub

Sub XAdd_Sub(SubNm$)
Md_XApp_Lines CurMd, QQ_Fmt("Sub ?()|End Sub", SubNm)
MdMth_XGo CurMd, SubNm
End Sub

Function Md_XAdd_OptExpLin(A As CodeModule) As CodeModule
A.InsertLines 1, "Option Explicit"
Set Md_XAdd_OptExpLin = A
End Function

Function Pj_XAdd_Cls(A As VBProject, Nm$) As CodeModule
Set Pj_XAdd_Cls = Md_XAdd_OptExpLin(Pj_XAdd_Cmp(A, Nm, vbext_ct_ClassModule).CodeModule)
End Function

Sub Pj_XAdd_ClsFmPj(A As VBProject, FmPj As VBProject, ClsNy0)
Dim I, ClsNy$(), ClsAy() As CodeModule
ClsNy = CvNy(ClsNy0)
For Each I In A
    Md_XCpy CvMd(I), A
Next
End Sub

Function Pj_XAdd_Cmp(A As VBProject, Nm, Ty As vbext_ComponentType) As VBComponent
Const CSub$ = CMod & "Pj_XAdd_Cmp"
Set Pj_XAdd_Cmp = Cmp_XSet_Nm(A.VBComponents.Add(Ty), Nm, CSub)
End Function

Function Pj_XAdd_CmpLines(A As VBProject, Nm, Ty As vbext_ComponentType, Lines$) As VBComponent
Dim O As VBComponent
Set O = Pj_XAdd_Cmp(A, Nm, Ty): If IsNothing(O) Then Stop
Md_XApp_Lines O.CodeModule, Lines
Set Pj_XAdd_CmpLines = O
End Function

Sub Pj_XAdd_MdPfx(A As VBProject, B As WhMd, MdPfx$)
Dim Md As CodeModule, M
For Each M In AyNz(Pj_MdAy(A, B))
    Set Md = M
    Md_XRen Md, MdPfx & Md_Nm(Md)
Next
End Sub

Function Pj_XAdd_Md(A As VBProject, Nm) As CodeModule
Set Pj_XAdd_Md = Md_XAdd_OptExpLin(Pj_XAdd_Cmp(A, Nm, vbext_ct_StdModule).CodeModule)
End Function

Sub Pj_XCrt_Cmp(A As VBProject, Nm, Ty As vbext_ComponentType)
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = Nm
End Sub

Sub Pj_XCrt_Md(A As VBProject, MdNm$)
Pj_XCrt_Cmp A, MdNm, vbext_ct_StdModule
End Sub

Function Pj_XEns_Cls(A As VBProject, ClsNm$) As CodeModule
Set Pj_XEns_Cls = Pj_XEns_Cmp(A, ClsNm, vbext_ct_ClassModule)
End Function

Function Pj_XEns_Cmp(A As VBProject, Nm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If Not Pj_XHas_CmpNm(A, Nm) Then
    Pj_XCrt_Cmp A, Nm, Ty
End If
Set Pj_XEns_Cmp = A.VBComponents(Nm).CodeModule
End Function

Function Pj_XEns_Md(A As VBProject, MdNm) As CodeModule
Set Pj_XEns_Md = Pj_XEns_Cmp(A, MdNm, vbext_ct_StdModule)
End Function

Function Pj_XEns_Std(A As VBProject, StdNm$) As CodeModule
Set Pj_XEns_Std = Pj_XEns_Cmp(A, StdNm, vbext_ct_StdModule)
End Function

Private Sub ZZ()
Dim A$
Dim B As CodeModule
Dim C As VBProject
Dim D As Variant
Dim E As vbext_ComponentType
Dim F As WhMd
XAdd_Cls A
XAdd_Fun A
XAdd_Mod A
XAdd_Sub A
Md_XAdd_OptExpLin B
Pj_XAdd_Cls C, A
Pj_XAdd_ClsFmPj C, C, D
Pj_XAdd_Cmp C, D, E
Pj_XAdd_CmpLines C, D, E, A
Pj_XAdd_MdPfx C, F, A
Pj_XAdd_Md C, D
Pj_XCrt_Cmp C, D, E
Pj_XCrt_Md C, A
Pj_XEns_Cls C, A
Pj_XEns_Cmp C, D, E
Pj_XEns_Md C, D
Pj_XEns_Std C, A
End Sub

Private Sub Z()
End Sub

Sub Pj_XAdd_MdNmDic(A As VBProject, MdNmDic As Dictionary)
Dim MdNm
For Each MdNm In MdNmDic.Keys
    Pj_XEns_Md A, MdNm
    Md_XApp_Lines Pj_Md(A, MdNm), MdNmDic(MdNm)
Next
End Sub


