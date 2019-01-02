Attribute VB_Name = "MIde__MTH_INF"
Option Explicit
Option Compare Database

Function CurVbe_MthDrs(Optional B As WhPjMth) As Drs
Set CurVbe_MthDrs = PjFfnAy_MthDrs(CurVbe_PjFfnAy, B)
End Function

Function Mth_MthDrs(A As CodeModule, Optional B As WhMth) As Drs
Set Mth_MthDrs = New_Drs(MthFny_OfMd, Md_MthDry(A, B))
End Function

Function Md_MthDry(A As CodeModule, Optional B As WhMth) As Variant()
Dim P As VBProject, Ffn$, Pj$, Ty$, Md$, Md_Ty$
Set P = Md_Pj(A)
Ffn$ = Pj_Ffn(P)
Pj = P.Name
Md_Ty = Md_TyStr(A)
Md = Md_Nm(A)
Md_MthDry = Dry_XIns_C4(Src_MthDry(Md_Src(A)), Ffn, Pj, Md_Ty, Md)
End Function

Function Md_MthLinDry(A As CodeModule, Optional B As WhMth) As Variant()
Dim Dry()
Dry = Src_MthLinDry(Md_Src(A), B)
Md_MthLinDry = Dry_XIns_CC(Dry, Md_MdShtTyNm(A), Md_Nm(A))
End Function

Function Md_MthLinDry_WrapPm(A As CodeModule, B As WhMth) As Variant()
Md_MthLinDry_WrapPm = Src_MthLinDry_WrapPm(Md_BdyLy(A), B)
End Function

Function CurVbe_Wb_MthInf(Optional A As WhPjMth) As Workbook
Wb_XVis PjFfnAy_Wb_MthInf(CurVbe_PjFfnAy, A)
End Function

Property Get CurVbeMthWs() As Worksheet
Set CurVbeMthWs = PjFfnAy_MthWs(CurVbe_PjFfnAy)
End Property

Function Fb_MthDrs(A, Optional B As WhPjMth) As Drs
If False Then
    Set Fb_MthDrs = Vbe_MthDrs(Fb_Acs(A).Vbe, B)
    Exit Function
End If
Dim Acs As New Access.Application
Debug.Print "FbMthDry: "; Now; " Start get Drs "; A; "==============="
Debug.Print "FbMthDry: "; Now; " Start open"
Set Acs = Fb_Acs(A)
Debug.Print "FbMthDry: "; Now; " Start get Drs "
Set Fb_MthDrs = Vbe_MthDrs(Acs.Vbe, B)
Debug.Print "FbMthDry: "; Now; " Start quit acs "
Acs.Quit acQuitSaveNone
Debug.Print "FbMthDry: "; Now; " acs is quit"
Set Acs = Nothing
Debug.Print "FbMthDry: "; Now; " acs is nothing"
End Function

Function XlsFxa_MthDrs(A As excel.Application, Fxa, Optional B As WhMdMth) As Drs
Dim Pj As VBProject
Set Pj = XlsFxa_Pj(A, Fxa)
If IsNothing(Pj) Then
    Fxa_XOpn Fxa
    Set Pj = VbePjFfn_Pj(CurXls.Vbe, Fxa)
    If IsNothing(Pj) Then Stop
End If
Set XlsFxa_MthDrs = Pj_MthDrs(Pj, B)
End Function

Function PjFfnAy_Wb_MthInf(PjFfnAy$(), Optional B As WhPjMth) As Workbook
Set PjFfnAy_Wb_MthInf = Wb_Fmt(Ws_Wb(PjFfnAy_MthWs(PjFfnAy, B)))
End Function

Function PjFfnAy_MthWs(PjFfnAy, Optional B As WhPjMth) As Worksheet
Dim O As Drs
Set O = PjFfnAy_MthDrs(PjFfnAy, B)
Set O = Drs_XAdd_ValIdCol(O, "Nm", "Vbe_Mth")
Set O = Drs_XAdd_ValIdCol(O, "Lines", "Vbe")
'Set PjFfnAy_MthWs = WsSetCdNmAndLoNm(Drs_Ws(O), "Mth")
End Function

Sub PjFfn_XEns_MthCache(PjFfn)
Dim D1 As Date
Dim D2 As Date
    D1 = Pj_FfnPjDte(PjFfn)
    D2 = PjFfn_CacheDte_MthInf(PjFfn)
Select Case True
Case D1 = 0:  Stop
Case D2 = 0:
Case D1 = D2: Exit Sub
Case D2 > D1: Stop
End Select
Stop '
'Drs_XIup_Dbt Pj_MthDrs_FmLiv(Pj_Ffn), MthDb, "MthCache", QQ_Fmt("Pj_Ffn='?'", Pj_Ffn)
End Sub

Function PjFfn_CacheDte_MthInf(PjFfn) As Date
PjFfn_CacheDte_MthInf = Dbq_Val(MthDb, QQ_Fmt("Select PjDte from Mth where Pjffn='?'", PjFfn))
End Function

Function Pj_MthDrs_FM_CACHE(PjFfn, Optional B As WhMdMth) As Drs
Dim Sql$: Sql = QQ_Fmt("Select * from MthCache where PjFfn='?'", PjFfn)
Set Pj_MthDrs_FM_CACHE = Dbq_Drs(MthDb, Sql)
End Function

Function Pj_MthDrs_FmLiv(PjFfn) As Drs
Dim V As Vbe, A, P As VBProject, PjDte As Date
Set A = PjFfn_App(PjFfn)
Set V = A.Vbe
Set P = VbePjFfn_Pj(V, PjFfn)
Select Case True
Case IsFb(PjFfn):  PjDte = Acs_PjDte(CvAcs(A))
Case IsFxa(PjFfn): PjDte = FileDateTime(PjFfn)
Case Else: Stop
End Select
Set Pj_MthDrs_FmLiv = Drs_XAdd_Col(Pj_MthDrs(P), "PjDte", PjDte)

If IsFb(PjFfn) Then
    CvAcs(A).CloseCurrentDatabase
End If
End Function

Function PjFfn_MthDrs(PjFfn, Optional B As WhMdMth) As Drs
PjFfn_XEns_MthCache PjFfn
Set PjFfn_MthDrs = Pj_MthDrs_FM_CACHE(PjFfn, B)
End Function
Function CurPj_MthDrs(Optional B As WhMdMth) As Drs
Set CurPj_MthDrs = Pj_MthDrs(CurPj, B)
End Function
Function Pj_MthDrs(A As VBProject, Optional B As WhMdMth) As Drs
Dim O As Drs
Set O = New_Drs(MthFny_OfMd, Pj_MthDry(A, B))
Set O = Drs_XAdd_ValIdCol(O, "Lines", "Pj")
Set O = Drs_XAdd_ValIdCol(O, "Nm", "PjMth")
Set Pj_MthDrs = O
End Function
Function Pj_MthLinDt(A As VBProject, B As WhMdMth) As Dt
Dim DtNm$
    DtNm = "Pj-MthLin-" & Pj_Nm(A)
Set Pj_MthLinDt = New_Dt(DtNm, MthFny_OfMd, Pj_MthLinDry(A, B))
End Function

Function CurPj_MthLinDt(Optional B As WhMdMth) As Dt
Set CurPj_MthLinDt = Pj_MthLinDt(CurPj, B)
End Function
Function CurPj_MthLinDrs(Optional B As WhMdMth) As Drs
Set CurPj_MthLinDrs = Pj_MthLinDrs(CurPj, B)
End Function

Function CurPj_MthLinWs(Optional B As WhMdMth) As Worksheet
Set CurPj_MthLinWs = Pj_MthLinWs(CurPj, B)
End Function

Sub CurPj_MthLinDrs_XBrw(Optional B As WhMdMth)
Drs_XBrw CurPj_MthLinDrs(B)
End Sub

Sub CurPj_MthLinWs_XBrw(Optional B As WhMdMth)
Ws_XVis CurPj_MthLinWs(B)
End Sub

Function Pj_MthLinDrs(A As VBProject, Optional B As WhMdMth) As Drs
Set Pj_MthLinDrs = New_Drs(MthFny_OfMd, Pj_MthLinDry(A, B))
End Function

Function Pj_MthLinWs(A As VBProject, Optional B As WhMdMth) As Worksheet
Set Pj_MthLinWs = Sq_Ws(Pj_MthLinSq(A, B))
End Function
Function Pj_MthLinSq(A As VBProject, Optional B As WhMdMth) As Variant()
Dim Dry() ' with hdr
Dry = AyIns(Pj_MthLinDry(A, B), MthFny_OfMd)
Pj_MthLinSq = Dry_Sq(Dry)
End Function

Function Pj_MthLinDry(A As VBProject, Optional B As WhMdMth) As Variant()
Dim M
For Each M In AyNz(Pj_MdAy(A, WhMdMth_WhMd(B)))
    PushIAy Pj_MthLinDry, Md_MthLinDry(CvMd(M), WhMdMth_WhMth(B))
Next
End Function

Function Pj_MthDry(A As VBProject, Optional B As WhMdMth) As Variant()
Dim M
For Each M In AyNz(Pj_MdAy(A, WhMdMth_WhMd(B)))
    PushIAy Pj_MthDry, Md_MthDry(CvMd(M), WhMdMth_WhMth(B))
Next
End Function

Function Pj_MthWs(A As VBProject, Optional B As WhMdMth) As Worksheet
Set Pj_MthWs = Drs_Ws(Pj_MthDrs(A, B))
End Function
Private Sub Z_PjFfnAy_Wb()
Wb_XVis PjFfnAy_Wb_MthInf(CurVbe_PjFfnAy, New_WhPjMth(MdMth:=New_WhMdMth(WhMd("Std"))))
End Sub

Private Sub Z_CurVbe_Wb_MthInf()
Wb_XVis CurVbe_Wb_MthInf
End Sub

Private Sub Z_Mth_MthDrs()
Drs_XBrw Mth_MthDrs(CurMd)
End Sub

Private Sub Z_Wb_Fmt()
Dim Wb As Workbook
Const Fx$ = "C:\Users\user\Desktop\Vba-Lib-1\Mth.xlsx"
Wb_Fmt Wb_XVis(Fx_Wb(Fx))
Stop
End Sub


Function PjFfnAy_MthDrs(PjFfnAy, Optional B As WhPjMth) As Drs
Dim I
For Each I In PjFfnAy
    Drs_Push PjFfnAy_MthDrs, PjFfn_MthDrs(I, B)
Next
End Function

Private Sub Z_Pj_MthDrs_FmLiv()
Dim A As Drs, A1$
A1 = CurVbe_PjFfnAy()(0)
Set A = Pj_MthDrs_FmLiv(A1)
Ws_XVis Drs_Ws(A)
End Sub

Function Vbe_MthDrs(A As Vbe, Optional B As WhPjMth) As Drs
Dim P
For Each P In AyNz(Vbe_PjAy(A, WhPjMth_WhNm(B)))
    Drs_Push Vbe_MthDrs, Pj_MthDrs(CvPj(P), WhPjMth_MdMth(B))
Next
End Function

Function Vbe_MthWs(A As Vbe, Optional B As WhPjMth) As Worksheet
Set Vbe_MthWs = Drs_Ws(Vbe_MthDrs(A, B))
End Function

Function Lin_MthLinDr(A) As Variant()
Dim L$, ShtMdy$, ShtTy$, Nm$, Prm$, Ret$, TopRmk$, LinRmk$
L = A
ShtMdy = XShf_MthShtMdy(L)
ShtTy = XShf_MthShtTy(L): If ShtTy = "" Then Exit Function
Nm = XShf_Nm(L)
Ret = XShf_MthSfx(L)
Prm = XShf_BktStr(L)
If XShf_X(L, "As") Then
    If Ret <> "" Then Stop
    Ret = XShf_Term(L)
End If
If XShf_Pfx(L, "'") Then
    LinRmk = L
End If
Lin_MthLinDr = Array(ShtMdy, ShtTy, Nm, Ret, Prm, LinRmk)
End Function

Function Lin_MthLinDr_WrapPm(A) As Variant()
Dim O()
O = Lin_MthLinDr(A)
If Sz(O) = 0 Then Exit Function
O(3) = Ay_XAdd_CommaSpcSfxExlLas(AyTrim(SplitComma(O(3))))
Lin_MthLinDr_WrapPm = O
End Function

Function Wb_Fmt(A As Workbook) As Workbook
Dim Ws As Worksheet, Lo As ListObject
Set Ws = Wb_Ws_BY_CD_NM(A, "MthLoc"): If IsNothing(Ws) Then Stop
Set Lo = Ws_Lo(Ws, "T_MthLoc"): If IsNothing(Lo) Then Stop
Dim Ws1 As Worksheet:  GoSub X_Ws1
Dim Pt1 As PivotTable: GoSub X_Pt1
Dim Lo1 As ListObject: GoSub X_Lo1
Dim Pt2 As PivotTable: GoSub X_Pt2
Dim Lo2 As ListObject: GoSub X_Lo2
Ws1.Outline.ShowLevels , 1
Set Wb_Fmt = Ws_Wb(Ws)
Exit Function
X_Ws1:
    Set Ws1 = Wb_XAdd_Ws(Ws_Wb(Ws))
    Ws1.Outline.SummaryColumn = xlSummaryOnLeft
    Ws1.Outline.SummaryRow = xlSummaryBelow
    Return
X_Pt1:
    Set Pt1 = Lo_Pt(Lo, Ws_A1(Ws1), "Md_Ty Nm VbeLinesId Lines", "Pj")
    PtSetRowssOutLin Pt1, "Lines"
    PtSetRowssColWdt Pt1, "VbeLinesId", 12
    PtSetRowssColWdt Pt1, "Nm", 30
    PtSetRowssRepeatLbl Pt1, "Md_Ty Nm"
    Return
X_Lo1:
    Set Lo1 = PtCpyToLo(Pt1, Ws1.Range("G1"))
    Lo_XSet_Nm Lo1, "T_MthLines"
    Lc_XSet_Wdt Lo1, "Nm", 30
    Lc_XSet_Wdt Lo1, "Lines", 100
    Lc_XSet_Lvl Lo1, "Lines"
    
    Return
X_Pt2:
    Set Pt2 = Lo_Pt(Lo1, Ws1.Range("M1"), "Md_Ty Nm", "Lines")
    PtSetRowssRepeatLbl Pt2, "Md_Ty"
    Return
X_Lo2:
    Set Lo2 = PtCpyToLo(Pt2, Ws1.Range("Q1"))
    Lo_XSet_Nm Lo2, "T_UsrEdtMthLoc"
    Return
Set Wb_Fmt = A
End Function

Function Pj_MthLinDry_WrapPm(A As VBProject, Optional B As WhMdMth) As Variant()
Dim M, WhMth As WhMth
Set WhMth = WhMdMth_WhMth(B)
For Each M In AyNz(Pj_MdAy(A, WhMdMth_WhMd(B)))
    PushIAy Pj_MthLinDry_WrapPm, Md_MthLinDry_WrapPm(CvMd(M), WhMth)
Next
End Function

Function Pj_Sq_MTH_KEY(A As VBProject) As Variant()
Pj_Sq_MTH_KEY = MthKy_Sq(Pj_MthKy(A, True))
End Function

Function Pj_Ws_MTH_KEY(A As CodeModule) As Worksheet
Set Pj_Ws_MTH_KEY = Ws_XVis(Sq_Ws(Pj_Sq_MTH_KEY(A)))
End Function

Function Src_DclDr(A$()) As Variant()
Dim Dcl$
Dcl = Src_DclLines(A): If Dcl = "" Then Exit Function
Dim Cnt%
Cnt = LinCnt(Dcl)
Const FF = "Ty Nm Cnt Lines"
Dim Vy(): Vy = Array("Dcl", "*Dcl", Cnt, Dcl)
Src_DclDr = VyDr(Vy, FF, MthFny_OfSrc)
End Function

Function Src_MthDry(A$()) As Variant()
PushI_SomSz Src_MthDry, Src_DclDr(A)
Dim Ix
For Each Ix In AyNz(Src_MthIxAy(A))
    PushI Src_MthDry, SrcMthIx_MthDr(A, CLng(Ix))
Next
End Function

Function SrcMthIx_MthDr(A$(), MthIx&) As Variant()
Dim L$, Lines$, TopRmk$, Lno&, Cnt%
    'If A(MthIx) = "Private Sub ZZ_Mth_CxtFTNo _" Then Stop
    L = SrcIx_ContLin(A, MthIx)
    Lno = MthIx + 1
    Lines = SrcMthIx_MthLines(A, MthIx)
    Cnt = SubStrCnt(Lines, vbCrLf) + 1
    TopRmk = SrcMthIx_TopRmk(A, MthIx)
Dim Dr(): Dr = Lin_MthLinDr(L): If Sz(Dr) = 0 Then Stop
SrcMthIx_MthDr = Ay_XAdd_(Dr, Array(Lno, Cnt, Lines, TopRmk))
End Function

Property Get MthFny_OfSrc() As String()
MthFny_OfSrc = Ssl_Sy("Mdy Ty Nm Ret Prm LinRmk Lno Cnt Lines TopRmk")
End Property

Property Get MthFny_OfPj() As String()
MthFny_OfPj = Ssl_Sy("PjFfn PjNm MdTy Md Mdy Ty Nm Ret Prm LinRmk Lno Cnt Lines TopRmk")
End Property

Property Get MthFny_OfMd() As String()
MthFny_OfMd = Ssl_Sy("MdTy Md Mdy Ty Nm Ret Prm LinRmk Lno Cnt Lines TopRmk")
End Property

Function VbeAy_MthDrs(A() As Vbe) As Drs
Dim I, R%, M As Drs
For Each I In AyNz(A)
    Set M = DrsInsCol(Vbe_MthDrs(CvVbe(I)), "Vbe", R)
    If R = 0 Then
        Set VbeAy_MthDrs = M
    Else
        Stop
        PushObj VbeAy_MthDrs, M
        Stop
    End If
    R = R + 1
    Debug.Print R; "<=== VbeAy_MthDrs"
Next
End Function

Function VbeAy_MthWs(A() As Vbe) As Worksheet
Set VbeAy_MthWs = Drs_Ws(VbeAy_MthDrs(A))
End Function

Private Property Get ZZVbeAy() As Vbe()
PushObj ZZVbeAy, CurVbe
Const Fb$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
PushObj ZZVbeAy, Fb_XOpn(Fb).Vbe
End Property

Private Sub Z_Pj_MthLinDry()
Dim A(): A = Pj_MthLinDry(CurPj)
Stop
End Sub

Private Sub Z_VbeAy_MthWs()
Ws_XVis VbeAy_MthWs(ZZVbeAy)
End Sub

Private Sub Z_Vbe_MthLinDry()
Brw Dry_Fmtss(Vbe_MthLinDry(CurVbe))
End Sub

Private Sub Z_Vbe_MthLinDryWP()
Brw Dry_FmtssWrp(Vbe_MthLinDryWP(CurVbe))
End Sub

Private Sub ZZ()
Dim A As WhPjMth
Dim B As Variant
Dim C As WhMdMth
Dim D As CodeModule
Dim E As WhMth
Dim F$
Dim G As Workbook
Dim H$()
Dim I As VBProject
Dim J&
Dim K() As Vbe
Dim L As Vbe
CurVbe_Wb_MthInf
Fb_MthDrs B, A
Lin_MthLinDr B
Lin_MthLinDr_WrapPm B
Mth_MthDrs D, E
Md_MthDry D, E
Wb_Fmt G
PjFfnAy_MthDrs B, A
PjFfnAy_MthWs B, A
Pj_MthDrs_FmLiv B
Pj_MthDrs I, C
Pj_MthDry I, C
Pj_Sq_MTH_KEY I
Src_DclDr H
Src_MthDry H
SrcMthIx_MthDr H, J
VbeAy_MthDrs K
VbeAy_MthWs K
Vbe_MthDrs L, A
Vbe_MthWs L, A
End Sub

Private Sub Z()
End Sub

