Attribute VB_Name = "MIde_Z_Md"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Z_Md."

Function Md_IsNoLin(A As CodeModule) As Boolean
Md_IsNoLin = A.CountOfLines = 0
End Function

Function Md_XHas_NoMth(A As CodeModule) As Boolean
Dim J&
For J = A.CountOfDeclarationLines + 1 To A.CountOfLines
    If Lin_IsMth(A.Lines(J, 1)) Then Exit Function
Next
Md_XHas_NoMth = True
End Function

Function CvMd(A) As CodeModule
Set CvMd = A
End Function

Function Md(MdDNm) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Sz(A1)
Case 1: Set Md = Pj_Md(CurPj, MdDNm)
Case 2: Set Md = Pj_Md(Pj(A1(0)), A1(1))
Case Else: XThw CSub, "[MdDNm] should be XXX.XXX or XXX", "Md", MdDNm
End Select
End Function

Function Md_SummaryLin$(A As CodeModule)
Md_SummaryLin = Md_Nm(A) & " " & Md_NMth(A) & " " & A.CountOfLines
End Function

Function MdAy_CmpAy(A() As CodeModule) As VBComponent()
Dim I
For Each I In AyNz(A)
    PushObj MdAy_CmpAy, Md_Cmp(CvMd(I))
Next
End Function

Function MdAy_XWh_InTy(A() As CodeModule, WhInTyAy0$) As CodeModule()
Dim TyAy() As vbext_ComponentType, Md
TyAy = CvWhCmpTy(WhInTyAy0)
Dim O() As CodeModule
For Each Md In A
    If Ay_XHas(TyAy, CvMd(Md).Parent.Type) Then PushObj O, Md
Next
MdAy_XWh_InTy = O
End Function

Function MdAy_XWh_Mdy(A() As CodeModule, CmpTyAy0) As CodeModule()
'MdAy_XWh_Mdy = Ay_XWh_PredXP(A, "Md_IsInCmpAy", CvCmpTyAy(CmpTyAy0))
End Function

Function Md_Nm$(A As CodeModule)
Md_Nm = A.Parent.Name
End Function

Function Md_CanHasCd(A As CodeModule) As Boolean
Select Case Md_Ty(A)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    Md_CanHasCd = True
End Select
End Function

Function Md_Cmp(A As CodeModule) As VBComponent
Set Md_Cmp = A.Parent
End Function

Function Md_CmpTy(A As CodeModule) As vbext_ComponentType
Md_CmpTy = A.Parent.Type
End Function

Function Md_DNm$(A As CodeModule)
Md_DNm = Mth_PjNm(A) & "." & Md_Nm(A)
End Function

Function Md_Fn$(A As CodeModule)
Md_Fn = Md_Nm(A) & Md_SrcExt(A)
End Function

Function Md_IsCls(A As CodeModule) As Boolean
Md_IsCls = A.Parent.Type = vbext_ct_ClassModule
End Function

Function Md_IsStd(A As CodeModule) As Boolean
Md_IsStd = A.Parent.Type = vbext_ct_StdModule
End Function

Function Md_MdShtTyNm$(A As CodeModule)
Md_MdShtTyNm = CmpTy_ShtNm(Md_CmpTy(A))
End Function

Function Md_MthBrkAy(A As CodeModule, Optional B As WhMth) As Variant()
Md_MthBrkAy = Src_MthNmDry(Md_Src(A), B)
End Function

Function Md_NTy%(A As CodeModule)
Md_NTy = Src_NTy(Md_DclLy(A))
End Function

Function Md_Pj(A As CodeModule) As VBProject
Set Md_Pj = A.Parent.Collection.Parent
End Function

Function Md_ResLy(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Const CSub$ = CMod & "Md_ResLy"
Dim Z$()
    Z = MdMthBdyLy(A, ResPfx & ResNm)
    If Sz(Z) = 0 Then
        XThw CSub, "MthNm not found", "MthNm Md ResNm ResPfx", ResPfx & ResNm, Md_Nm(A), ResNm, ResPfx
    End If
    Z = Ay_XExl_FstEle(Z)
    Z = Ay_XExl_LasEle(Z)
Md_ResLy = Ay_XRmv_FstChr(Z)
End Function

Function Md_ResStr$(A As CodeModule, ResNm$)
Md_ResStr = JnCrLf(Md_ResLy(A, ResNm))
End Function

Function Md_Src(A As CodeModule) As String()
Md_Src = Md_Ly(A)
End Function

Function Md_SrcLines$(A As CodeModule)
Md_SrcLines = JnCrLf(Md_Src(A))
End Function

Function Md_Sz&(A As CodeModule)
Md_Sz = Len(Md_SrcLines(A))
End Function

Function Md_TopRmkMthLinesAy(A As CodeModule) As String()
Md_TopRmkMthLinesAy = Src_MthNmDicTopRmkMthLinesAy(Md_MthDic(A))
End Function

Function Md_Ty(A As CodeModule) As vbext_ComponentType
Md_Ty = A.Parent.Type
End Function

Function Md_TyStr$(A As CodeModule)
Md_TyStr = CmpTy_ShtNm(A.Parent.Type)
End Function

Sub Md_XCls_Win(A As CodeModule)
A.CodePane.Window.Close
End Sub

Sub Md_XCmp(A As CodeModule, B As CodeModule)
Dim A1 As Dictionary, B1 As Dictionary
    Set A1 = Md_MthDic(A)
    Set B1 = Md_MthDic(B)
Dic_Cmp_XBrw A1, B1, Md_DNm(A), Md_DNm(B)
End Sub

Sub Md_XRen_RmvNmPfx(A As CodeModule, Pfx$)
Dim Nm$: Nm = Md_Nm(A): If Not XHas_Pfx(Nm, Pfx) Then Exit Sub
Md_XRen A, XRmv_Pfx(Md_Nm(A), Pfx)
End Sub

Sub Md_XRmv_FmCnt(A As CodeModule, FmCnt As FmCnt)
With FmCnt
    If .Cnt = 0 Then Exit Sub
    A.DeleteLines .FmLno, .Cnt
End With
End Sub

Sub Md_XRmv_FmCntAy(A As CodeModule, B() As FmCnt)
If Not FmCntAy_IsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Sub Md_XRpl(A As CodeModule, NewMd_Lines$)
Md_XClr A
Md_XApp_Lines A, NewMd_Lines
End Sub

Sub Md_XRpl_Bdy(A As CodeModule, NewMdBdy$)
Md_XClr_Bdy A
Md_XApp_Lines A, NewMdBdy
End Sub

Sub Md_XRpl_Lin(A As CodeModule, Lno, NewLin$)
With A
    .DeleteLines Lno
    .InsertLines Lno, NewLin
End With
End Sub

Sub Md_XRpl_Ly(A As CodeModule, DclLy$())
Md_XRmv_Dcl A
A.InsertLines 1, JnCrLf(DclLy)
End Sub

Sub Md_XShw(A As CodeModule)
A.CodePane.Show
End Sub
Function Md_PjNm$(A As CodeModule)
Md_PjNm = Md_Pj(A).Name
End Function


Sub XSrt()
Md_XSrt CurMd
End Sub

Private Property Get ZZMd() As CodeModule
Set ZZMd = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Property

Private Sub ZZ_MdDrs()
'Drs_XBrw MdDrs(Md("IdeFeature_XEns_ZZ_AsPrivate"))
End Sub

Private Sub ZZ_MdMthNm_Lno()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = Md_MthNy(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MdMthNm_Lno(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Sz(Ny), "Z_MdMthNm_Lno"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
Ay_XBrw O
End Sub

Private Sub ZZ_Md_XSrt()
Dim Md As CodeModule
GoSub X0
Exit Sub
X0:
    Dim I
'    For Each I In Pj_MdAy(CurPj_X)
        Set Md = I
        If Md_Nm(Md) = "Str_" Then
            GoSub Ass
        End If
'    Next
    Return
X1:

    Return
Ass:
    Debug.Print Md_Nm(Md); vbTab;
    Dim BefSrt$(), AftSrt$()
    BefSrt = Md_Ly(Md)
    AftSrt = SplitCrLf(Md_SrtLines(Md))
    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
        Debug.Print "Is Same of before and after sorting ......"
        Return
    End If
    If Sz(AftSrt) <> 0 Then
        If Ay_LasEle(AftSrt) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & Md_Nm(Md) & "=====")
            Ay_XBrw AyAp_XAdd(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = AyMinus(BefSrt, AftSrt)
    B = AyMinus(AftSrt, BefSrt)
    Debug.Print
    If Sz(A) = 0 And Sz(B) = 0 Then Return
    If Sz(Ay_XExl_EmpEle(A)) <> 0 Then
        Debug.Print "Sz(A)=" & Sz(A)
        Ay_XBrw A
        Stop
    End If
    If Sz(Ay_XExl_EmpEle(B)) <> 0 Then
        Debug.Print "Sz(B)=" & Sz(B)
        Ay_XBrw B
        Stop
    End If
    Return
End Sub

Private Sub ZZ_Md_SrtLines()
StrBrw Md_SrtLines(CurMd)
End Sub

Private Sub Z_MdEnmMbrCnt()
Ass MdEnmMbrCnt(Md("Ide"), "AA") = 1
End Sub

Private Sub Z_MdMthNm_MthLines()
Debug.Print Len(MdMthNm_MthLines(CurMd, "MdMthNm_Lines"))
Debug.Print MdMthNm_MthLines(CurMd, "MdMthNm_Lines")
End Sub

Private Sub Z_MdMthBdyLy()
Debug.Print Len(MdMthNm_MthLines(CurMd, "MdMthNm_Lines"))
Debug.Print MdMthNm_MthLines(CurMd, "MdMthNm_Lines")
End Sub

Private Sub Z_MdMthLno_MthNLin()
Dim O$()
    Dim J%, M, L&, E&, A As CodeModule, Ny$()
    Set A = Md("Fct")
    Ny = Md_MthNy(A)
    For Each M In Ny
        DoEvents
        L = MdMthNm_Lno(A, CStr(M))
        E = MdMthLno_MthNLin(A, L) + L - 1
        Push O, Format(L, "0000 ") & A.Lines(L, 1)
        Push O, Format(E, "0000 ") & A.Lines(E, 1)
    Next
Ay_XBrw O
End Sub

Private Sub Z_MdRmvPrpOnEr()
MdRmvPrpOnEr ZZMd
End Sub

Private Sub Z_Md_XEns_PrpOnEr()
Md_XEns_PrpOnEr ZZMd
End Sub

Private Sub Z_Md_Ly()
Ay_XBrw Md_Ly(CurMd)
End Sub

''======================================================================================
Private Sub Z_Md_MthDDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
Brw Md_MthNy(Md1)
Brw Md_MthDDNy(Md1)
End Sub

Private Sub Z_Md_TopRmkMthLinesAy()
Brw Jn(Md_TopRmkMthLinesAy(CurMd), vbCrLf & "-----------------------------------------------" & vbCrLf)
End Sub

Private Sub Z_Md_XApp_Lines()
Const MdNm$ = "Module1"
Md_XApp_Lines CurMd, "'aa"
End Sub

Private Sub Z_Md_XCpy()
Dim A As CodeModule, ToPj As VBProject
'
Set ToPj = CurPj
Set A = Md("QDta.Dt")
GoSub Tst
Exit Sub
Tst:
    Dim N$
    N = Md_Nm(A)
    Stop
    Md_XCpy A, ToPj   '<====
    Ass Pj_XHas_MdNm(ToPj, N) = True
    Pj_XDlt_Md ToPj, N
    Return
End Sub

Private Sub Z_Md_XRmv_FmCntAy()
Dim A() As FmCnt
A = MdMdMthNm_FmCntAy(Md("Md_"), "XXX")
Md_XRmv_FmCntAy Md("Md_"), A
End Sub

Private Sub Z_Md_XTrim_END()
Dim M As CodeModule: Set M = Md("ZZModule")
Md_XApp_Lines M, "  "
Md_XApp_Lines M, "  "
Md_XApp_Lines M, "  "
Md_XApp_Lines M, "  "
Ass M.CountOfLines = 15
End Sub

Private Sub ZZ()
Dim A As CodeModule
Dim B
Dim C() As CodeModule
Dim D$
Dim E As WhMth
Dim F As FmCnt
Dim G() As FmCnt
Dim H$()
Dim XX
Md_IsNoLin A
Md_XHas_NoMth A
CvMd B
Md B
MdAy_CmpAy C
MdAy_XWh_InTy C, D
MdAy_XWh_Mdy C, B
Md_Nm A
Md_Cmp A
Md_CmpTy A
Md_DNm A
Md_Fn A
Md_IsCls A
Md_IsStd A
Md_MdShtTyNm A
Md_MthBrkAy A, E
Md_NTy A
Md_Pj A
Md_ResLy A, D, D
Md_ResStr A, D
Md_Src A
Md_SrcLines A
Md_Sz A
Md_TopRmkMthLinesAy A
Md_Ty A
Md_TyStr A
Md_XCls_Win A
Md_XCmp A, A
Md_XRen_RmvNmPfx A, D
Md_XRmv_FmCnt A, F
Md_XRmv_FmCntAy A, G
Md_XRpl A, D
Md_XRpl_Bdy A, D
Md_XRpl_Lin A, B, D
Md_XRpl_Ly A, H
Md_XShw A
Mth_PjNm A
XSrt
End Sub

Private Sub Z()
Z_MdEnmMbrCnt
Z_MdMthNm_MthLines
Z_MdMthBdyLy
Z_MdMthLno_MthNLin
Z_MdRmvPrpOnEr
Z_Md_XEns_PrpOnEr
Z_Md_Ly
Z_Md_MthDDNy
Z_Md_TopRmkMthLinesAy
Z_Md_XApp_Lines
Z_Md_XCpy
Z_Md_XRmv_FmCntAy
Z_Md_XTrim_END
End Sub
