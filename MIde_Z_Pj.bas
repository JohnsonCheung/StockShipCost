Attribute VB_Name = "MIde_Z_Pj"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Z_Pj."

Sub XAss_CompileBtn(Pj_Nm$)
If CompileBtn.Caption <> "Compi&le " & Pj_Nm Then Stop
End Sub

Property Get CurPj_NCls%()
CurPj_NCls = Pj_NCls(CurPj)
End Property
Property Get CurPj_NLin&()
CurPj_NLin = Pj_NLin(CurPj)
End Property
Function Pj_NLin&(A As VBProject)
Dim O&, I
For Each I In Pj_MdAy(A)
    O = O + CvMd(I).CountOfLines
Next
Pj_NLin = O
End Function
Property Get CurPj_NCmp%()
CurPj_NCmp = CurPj.VBComponents.Count
End Property

Property Get CurPj_NMd%()
CurPj_NMd = Pj_NMd(CurPj)
End Property

Property Get CurPj_SrcLinCnt&()
CurPj_SrcLinCnt = Pj_SrcLinCnt(CurPj)
End Property

Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function Is_PjNm(A) As Boolean
Is_PjNm = Ay_XHas(PjNy, A)
End Function

Function Pj(A) As VBProject
If A = "" Then Set Pj = CurPj: Exit Function
Set Pj = CurVbe.VBProjects(A)
End Function

Function Pj_Ffn$(A As VBProject)
On Error Resume Next
Pj_Ffn = A.FileName
End Function

Function PjFfn_App(PjFfn) ' Return either Xls.Application (CurXls) or Acs.Application (Function-static)
Static Y As New Access.Application
Select Case True
Case IsFxa(PjFfn): Fxa_XOpn PjFfn: Set PjFfn_App = CurXls
Case IsFb(PjFfn): Y.OpenCurrentDatabase PjFfn: Set PjFfn_App = Y
Case Else: Stop
End Select
End Function

Function Pj_Fn$(A As VBProject)
Pj_Fn = Ffn_Fn(Pj_Ffn(A))
End Function

Function Pj_FunPfxAy(A As VBProject) As String()
Dim Ay() As CodeModule: Ay = Pj_MdAy(A)
Dim Ay1(): Ay1 = AyMap(Ay, "MdFunPfx")
Pj_FunPfxAy = AyFlat(Ay1)
End Function

Function Pj_IsUsrLib(A As VBProject) As Boolean
Pj_IsUsrLib = Pj_IsFxa(A)
End Function

Function Pj_Md_MAX_LIN_CNT(A As VBProject) As CodeModule
Dim O As CodeModule, M&, N&, I
For Each I In Pj_MdAy(A)
    N = CvMd(I).CountOfLines
    If N > M Then
        M = N
        Set O = I
    End If
Next
Set Pj_Md_MAX_LIN_CNT = O
End Function
Function Pj_Md(A As VBProject, Nm) As CodeModule
Set Pj_Md = PjCmp(A, Nm).CodeModule
End Function
Function Pj_NCls(A As VBProject)
Pj_NCls = Pj_NCmp_TY(A, vbext_ct_ClassModule)
End Function

Function Pj_NCmp%(A As VBProject)
Pj_NCmp = A.VBComponents.Count
End Function

Function Pj_NCmp_OTH%(A As VBProject)
Pj_NCmp_OTH = Pj_NCmp(A) - Pj_NCls(A) - Pj_Nmod(A)
End Function

Function Pj_NCmp_TY%(A As VBProject, Ty As vbext_ComponentType)
Dim O%, C As VBComponent
For Each C In A.VBComponents
    If C.Type = Ty Then O = O + 1
Next
Pj_NCmp_TY = O
End Function

Function Pj_NMd%(A As VBProject)
Pj_NMd = Pj_NCmp_TY(A, vbext_ct_StdModule)
End Function

Function Pj_Nm$(A As VBProject)
Pj_Nm = A.Name
End Function

Function Pj_Nmod%(A As VBProject)
Pj_Nmod = Pj_NCmp_TY(A, vbext_ct_StdModule)
End Function

Function Pj_Pth$(A As VBProject)
Pj_Pth = Ffn_Pth(A.FileName)
End Function

Function Pj_SrcLinCnt&(A As VBProject)
Dim O&, C As VBComponent
For Each C In A.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
Pj_SrcLinCnt = O
End Function

Function Pj_SrcPth$(A As VBProject)
Dim P$:
P = Ffn_Pth(A.FileName) & "Src\" & Ffn_Fn(A.FileName) & "\"
Pj_SrcPth = Pth_XEnsAll(P)
End Function

Sub Pj_SrcPth_XBrw(A As VBProject)
Pth_XBrw Pj_SrcPth(A)
End Sub

Function Pj_Tim(A As VBProject) As Date
Pj_Tim = Ffn_Tim(Pj_Ffn(A))
End Function

Function Pj_Vbe(A As VBProject) As Vbe
Set Pj_Vbe = A.Collection.Vbe
End Function

Sub Pj_XCompile(A As VBProject)
Pj_XGo A
XAss_CompileBtn Pj_Nm(A)
With CompileBtn
    If .Enabled Then
        .Execute
        Debug.Print Pj_Nm(A), "<--- Compiled"
    Else
        Debug.Print Pj_Nm(A), "already Compiled"
    End If
End With
TileVBtn.Execute
SavBtn.Execute
End Sub

Sub Pj_XCpy_ToSrcPth(A As VBProject)
Ffn_XCpy_ToPth A.FileName, Pj_SrcPth(A), OvrWrt:=True
End Sub

Sub Pj_XCpy_ToSrcPth_Pth(A As VBProject)
Ffn_XCpy_ToPth A.FileName, Pj_SrcPth(A), OvrWrt:=True
End Sub

Sub Pj_XImp_SrcFfn(A As VBProject, SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Sub Pj_XSav(A As VBProject)
'If XTak_FstChr(Pj_Nm(A)) <> "Q" Then
'    Debug.Print "Pj_XSav: Project Name begin with Q, it is not saved: "; Pj_Nm(A)
'    Exit Sub
'End If
Const CSub$ = CMod & "Pj_XSav"
If A.Saved Then
    Debug.Print QQ_Fmt("Pj_XSav: Pj(?) is already saved", A.Name)
    Exit Sub
End If
Dim Fn$: Fn = Pj_Fn(A)
If Fn = "" Then
    XThw CSub, "Pj file name is blank.  The pj needs to saved first in order to have a pj file name", "Pj", Pj_Nm(A)
End If
Pj_XAct A
    Dim O1&, O2&
        O1 = 1
        O2 = 2
        O1 = ObjPtr(Vbe.ActiveVBProject)
        O2 = ObjPtr(A)
    If O1 <> O2 Then
        XThw CSub, "Error in Pj_XAct, the ObjPtr are diff", "ActPj-ObjPtr Given-Pj-ObjPtr", O1, O2
        Exit Sub
    End If
Dim B As CommandBarButton
    Set B = Vbe_SavBtn(Pj_Vbe(A))
    If Not Str_IsEq(B.Caption, "&Save " & Fn) Then Stop
B.Execute
If A.Saved Then
    Debug.Print QQ_Fmt("Pj_XSav: Pj(?) is saved <---------------", A.Name)
Else
    Debug.Print QQ_Fmt("Pj_XSav: Pj(?) cannot be saved for unknown reason <=================================", A.Name)
End If
End Sub

Private Sub ZZ_Pj_XHas_MdNm()
Ass Pj_XHas_MdNm(CurPj, "Drs") = False
Ass Pj_XHas_MdNm(CurPj, "A__Tool") = True
End Sub

Private Sub ZZ_Pj_XSav()
Pj_XSav CurPj
End Sub

Private Sub Z_Pj_XAdd_MdNmDic()
Dim MdNmDic As New Dictionary
Dim ToPj As VBProject: Set ToPj = TmpFxaPj
Pj_XAdd_MdNmDic ToPj, MdNmDic
End Sub

Private Sub Z_Pj_XCompile()
Pj_XCompile CurPj
End Sub

Private Sub ZZ()
Dim A$
Dim B As Variant
Dim C As VBProject
Dim D As Dictionary
Dim E As vbext_ComponentType
XAss_CompileBtn A
Dmp CurPj_NCmp
Dmp CurPj_NMd
CvPj B
Is_PjNm B
Pj B
Dmp PjNy
Pj_Ffn C
PjFfn_App B
Pj_Fn C
Pj_FunPfxAy C
Pj_IsUsrLib C
Pj_Md C, B
Pj_XAdd_MdNmDic C, D
Pj_NCls C
Pj_NCmp C
Pj_NCmp_OTH C
Pj_NCmp_TY C, E
Pj_NMd C
Pj_Nm C
Pj_Nmod C
Pj_Pth C
Pj_SrcLinCnt C
Pj_SrcPth C
Pj_SrcPth_XBrw C
Pj_Tim C
Pj_Vbe C
Pj_XCompile C
Pj_XCpy_ToSrcPth C
Pj_XCpy_ToSrcPth_Pth C
Pj_XImp_SrcFfn C, B
Pj_XSav C
End Sub

Private Sub Z()
Z_Pj_XAdd_MdNmDic
Z_Pj_XCompile
End Sub
