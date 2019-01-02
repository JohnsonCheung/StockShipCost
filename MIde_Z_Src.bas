Attribute VB_Name = "MIde_Z_Src"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Z_Src."
Property Get CurSrc() As String()
CurSrc = Md_Src(CurMd)
End Property


Private Sub ZZ_Src_Dcl()
StrBrw Src_DclLy(ZZSrc)
End Sub

Private Sub ZZ_Src_FstMthIx()
Dim Act%
Act = Src_FstMthIx(ZZSrc)
Ass Act = 2
End Sub

Private Sub ZZ_SrcMthNm_MthLines()
Dim Src$(): Src = ZZSrc
Dim MthNm$: MthNm = "A"
Dim Act$()
Act = SrcMthNm_MthLines(Src, MthNm)
End Sub

Private Sub ZZ_SrcMthIx_MthIxTopRmkFm()
Dim ODry()
    Dim Src$(): Src = Md_Src(Md("IdeSrcLin"))
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, L
    For Each L In Src
        IsMth = ""
        RmkLx = ""
        If Lin_IsMth(L) Then
            IsMth = "*Mth"
            RmkLx = SrcMthIx_MthIxTopRmkFm(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, L)
        Push ODry, Dr
        Lx = Lx + 1
    Next
Drs_XBrw New_Drs("Mth RmkLx Lin", ODry)
End Sub

Private Sub ZZ_Src_MthIxAy()
Dim IxAy&(): IxAy = Src_MthIxAy(CurSrc)
Dim Ay$(): Ay = Ay_XWh_IxAy(CurSrc, IxAy)
Dim Dry(): Dry = Ay_XZip(IxAy, Ay)
Dim O$()
O = Drs_Fmt(New_Drs("Ix Lin", Ay_XZip(IxAy, Ay)))
PushAy O, Drs_Fmt(AyDrs(CurSrc))
Ay_XBrw O
End Sub

Private Sub ZZ_Src_MthLxAy1()
Dim Src$(): Src = Md_Src(Md("DaoDb"))
Dim Ay$(): Ay = Ay_XWh_IxAy(Src, Src_MthIxAy(Src))
Ay_XBrw Ay
End Sub

Private Sub ZZ_Src_MthNy()
Dim Act$()
   Act = Src_MthNy(ZZSrc)
   Ay_XBrw Act
End Sub

Private Property Get ZZSrc() As String()
ZZSrc = Md_Src(Md("IdeSrc"))
End Property

Private Property Get ZZSrcLin$()
ZZSrcLin = "Private Sub Lin_IsMth()"
End Property
Private Sub Z_Src_MthNy()
Brw Src_MthNy(Md_Src(Md("AAAMod")))
End Sub
Function Src_MthNmDicTopRmkMthLinesAy(Src_MthNmDic As Dictionary) As String()
Dim L
For Each L In Src_MthNmDic.Items
    If XTak_FstChr(L) = "'" Then
        PushI Src_MthNmDicTopRmkMthLinesAy, L
    End If
Next
End Function

Function Src_MthKeyLinesDic1(A$()) As Dictionary
'To be delete
'Dim Ix, O As New Dictionary
'Src_MthNmDicAddDcl O, A
'For Each Ix In AyNz(SrcMthIx(A))
'    O.Add LinMthKey(A(Ix)), SrcMthIx_MthLines_WithTopRmk(A, Ix)
'Next
'Set SrcMthKeyLinesDic = O
End Function

Function MdNm_Src(MdNm$) As String()
MdNm_Src = Md_Ly(Md(MdNm))
End Function

Function Src_XAdd_Mth_IfNotExist(A$(), MthNm$, NewMthLy$()) As String()
If Src_XHas_MthNm(A, MthNm) Then
   Src_XAdd_Mth_IfNotExist = A
Else
   Src_XAdd_Mth_IfNotExist = AyAp_XAdd(A, NewMthLy)
End If
End Function

Function Src_BdyLines$(A$())
Src_BdyLines = JnCrLf(Src_BdyLy(A))
End Function

Function Src_BdyFmCnt(A$()) As FmCnt
Dim Lno&
Dim Cnt&
   Lno = Src_DclLinCnt(A) + 1
   Cnt = Sz(A) - Lno + 1
Set Src_BdyFmCnt = New_FmCnt(Lno, Cnt)
End Function

Function Src_BdyLy(A$()) As String()
Src_BdyLy = Ay_XWh_Fm(A, Src_DclLinCnt(A))
End Function

Function Src_CmpFmt(A1$(), A2$()) As String()
Dim D1 As Dictionary: Set D1 = Src_MthNmDic(A1)
Dim D2 As Dictionary: Set D2 = Src_MthNmDic(A2)
Src_CmpFmt = Dic_CmpFmt(D1, D2)
End Function

Function Src_MthNy_Dist(A$()) As String()
Dim O$(), I
If Sz(A) = 0 Then Exit Function
For Each I In A
   PushNonEmp O, Lin_MthNm(CStr(I))
Next
Src_MthNy_Dist = O
End Function

Function Src_XEns_Mth(T$(), MthNm$, NewMthLy$()) As String()
Src_XEns_Mth = Src_XAdd_Mth_IfNotExist(T, MthNm, NewMthLy)
End Function

Function Src_XHas_MthNm(A$(), MthNm, Optional WhMdy$, Optional WhKd$) As Boolean
Dim O&: O = SrcMthNm_MthIx(A, MthNm)
If O < 0 Then Exit Function
Dim Mdy$, Ty$, Nm$
Lin_MthNmBrkAsg A(O), Mdy, Ty, Nm
Src_XHas_MthNm = True
'If MthShtMdy_IsSel(Mdy, WhMdy) Then Exit Function
'If MthShtKd_IsSel(MthShtTy_MthShtKd(Ty), WhMthKd) Then Exit Function
Src_XHas_MthNm = False
End Function
Property Get SrcMthFny() As String()
Static X As Boolean, Y$()
If Not X Then
    X = True
    Y = SplitSpc("Md Lno Lin EnmNm IsBlank IsEmn IsMth Lin_IsPrp IsRmk IsTy Mdy MthNm MthTy NoMdy PrpTy TyNm")
End If
SrcMthFny = Y
End Property


Function SrcMthNm_MthLines(A$(), MthNm, Optional WithTopRmk As Boolean)
Dim I, O$()
For Each I In AyNz(SrcMthNm_MthIxAy(A, MthNm))
    PushI O, SrcMthIx_MthLines(A, CLng(I), WithTopRmk)
Next
SrcMthNm_MthLines = Join(O, vbCrLf & vbCrLf)
End Function

Private Sub Z_SrcMthNm_MthLines()
Dim MthNm$: MthNm = "A"
Dim Act$()
Act = SrcMthNm_MthLines(CurSrc, MthNm)
End Sub

Function SrcMthNm_FmCntAy(A$(), MthNm) As FmCnt()
Dim FmAy&(): FmAy = SrcMthNm_MthIxAy(A, MthNm)
Dim O() As FmCnt, J%
Dim ToIx&
Dim FTIx As FTIx
Dim FmCnt As FmCnt
For J = 0 To UB(FmAy)
   ToIx = SrcMthFmIx_MthToIx(A, FmAy(J))
   FTIx = FTIx(FmAy(J), ToIx)
   FmCnt = FTIx_FmCnt(FTIx)
   PushObj O, FmCnt
Next
SrcMthNm_FmCntAy = O
End Function

Sub SrcMthDr_XAsg(A, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AyAsg A, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Function Src_MthFTIxAy(A$()) As FTIx()
Dim F&(): F = Src_MthIxAy(A)
Dim U%: U = UB(F)
If U = -1 Then Exit Function
Dim O() As FTIx
ReDim O(U)
Dim J%
For J = 0 To U
    Set O(J) = New_FTIx(F(J), SrcMthIx_MthIxTo(A, F(J)))
Next
Src_MthFTIxAy = O
End Function
Function SrcMthNm_FTIxAy(A$(), MthNm) As FTIx()
Dim F&()
F = SrcMthNm_MthIxAy(A, MthNm): If Sz(F) <= 0 Then Exit Function
Dim O() As FTIx
ReDim O(UB(F))
Dim J%
For J = 0 To UB(F)
    Set O(J) = New_FTIx(F(J), SrcMthIx_MthIxTo(A, F(J)))
Next
SrcMthNm_FTIxAy = O
End Function

Function SrcMthIx_MthLines_WithTopRmk$(A$(), MthIx&)
Dim B$: B = SrcMthIx_TopRmk(A, MthIx)
Dim C$: C = SrcMthIx_MthLines(A, MthIx)
If B <> "" Then C = B & vbCrLf & C
SrcMthIx_MthLines_WithTopRmk = C
End Function
Function SrcMthIx_MthNLin&(A$(), MthIx&)
Dim ToIx&, O&
ToIx = SrcMthIx_MthIxTo(A, MthIx)
O = ToIx - MthIx + 1
If O < 0 Then Stop
SrcMthIx_MthNLin = O
End Function
Function SrcMthIx_MthLines$(A$(), MthIx&, Optional WithTopRmk As Boolean)
Dim L2&, Fm&
L2 = SrcMthIx_MthIxTo(A, MthIx)
If WithTopRmk Then
    Fm = SrcMthIx_MthIxTopRmkFm(A, MthIx)
Else
    Fm = MthIx
End If
SrcMthIx_MthLines = Join(SyWhFmTo(A, Fm, L2), vbCrLf)
End Function

Sub Src_MthNmDicAddDcl(A As Dictionary, Src$())
Dim Dcl$
Dcl = Src_DclLines(Src)
If Dcl = "" Then Exit Sub
A.Add "*Dcl", Dcl
End Sub

Private Sub Z_Src_MthNmDic()
Dic_XBrw Src_MthNmDic(Md_Src(Md("AAAMod")))
End Sub

Function Src_MthKy(A$()) As String()
Dim Ix
For Each Ix In AyNz(Src_MthIxAy(A))
    PushI Src_MthKy, Lin_MthSrtKey(A(Ix))
Next
End Function
Function Src_MthLinDry_WrapPm(A$(), B As WhMth) As Variant()
Dim L
For Each L In AyNz(A)
    PushI_SomSz Src_MthLinDry_WrapPm, Lin_MthLinDr_WrapPm(L)
Next
End Function

Function Src_MthLinesDic(A$(), Optional ExlDcl As Boolean) As Dictionary
Dim L&(): L = Src_MthIxAy(A)
Dim O As New Dictionary
    If Not ExlDcl Then O.Add "*Dcl", Src_DclLines(A)
    If Sz(L) = 0 Then GoTo X
    Dim MthNm$, Lin$, Lines$, Lx
    For Each Lx In L
        Lin = A(Lx)
        MthNm = Lin_MthNm(Lin):            If MthNm = "" Then Stop
        Lines = SrcMthIx_MthLines(A, CLng(Lx)): If Lines = "" Then Stop
        If O.Exists(MthNm) Then
            If Not Lin_IsPrp(Lin) Then Stop
            O(MthNm) = O(MthNm) & vbCrLf & vbCrLf & Lines
        Else
            O.Add MthNm, Lines
        End If
    Next
X:
Set Src_MthLinesDic = O
End Function

Function SrcMthFmIx_MthToIx&(A$(), MthFmIx)
Const CSub$ = CMod & "SrcMthFmIx_MthToIx"
Dim Lin$
   Lin = A(MthFmIx)

Dim E$
   E = MthLin_MthEndLin(Lin)
Dim O&
    For O = MthFmIx + 1 To UB(A)
        If XHas_Pfx(A(O), E) Then
            SrcMthFmIx_MthToIx = O
            Exit Function
        End If
    Next
XThw CSub, "No End-XXX line in Src", "MthFmIx MthLin Expected-End-XXX-Lin Src", MthFmIx, Lin, E, A
End Function

Function SrcMthMm_FmCntAy(A$(), MthNm) As FmCnt()
SrcMthMm_FmCntAy = FTIxAy_FmCntAy(SrcMthNm_FmCnt_Fst(A, MthNm))
End Function

Function SrcMthNm_FmCnt_Fst(A$(), MthNm) As FTIx()
Dim ToIx%, Fm
For Each Fm In AyNz(SrcMthNm_MthIx(A, MthNm))
    ToIx = SrcMthIx_MthIxTo(A, CLng(Fm))
    Push SrcMthNm_FmCnt_Fst, New_FTIx(Fm, ToIx)
Next
End Function

Private Sub Z_Src_MthDDNy()
Brw Src_MthDDNy(Md_Src(Md("AAAMod")))
End Sub

Function Src_MthDDNy(A$()) As String()
Dim L
For Each L In AyNz(A)
    PushNonBlankStr Src_MthDDNy, Lin_MthDDNm(CStr(L))
Next
End Function

Function Src_MthNmDry(A$(), Optional B As WhMth) As Variant()
Dim L
For Each L In AyNz(A)
    PushI_SomSz Src_MthNmDry, Lin_MthNmBrk(L, B)
Next
End Function

Function Src_NMth_Dist%(A$())
Src_NMth_Dist = Sz(Src_MthNy_Dist(A))
End Function

Function Src_NTy%(A$())
If Sz(A) = 0 Then Exit Function
Dim I, O%
For Each I In A
'   If SrcLin_IsTy(I) Then O = O + 1
Next
Src_NTy = O
End Function

Property Get SrcPth$()
Dim X As Boolean, Y$
If Not X Then
    X = True
    Y = CurDb_Pth & "Src\"
    Pth_XEns Y
End If
SrcPth = Y
End Property

Sub SrcPth_XBrw()
Pth_XBrw SrcPth
End Sub

Function Src_XRmv_Mth(A$(), MthNm) As String()
Dim FTIxAy() As FTIx
   FTIxAy = SrcMthNm_FmCnt_FstIxAy(A, MthNm)
Dim O$()
   O = A
   Dim J%
   For J = UB(FTIxAy) To 0 Step -1
       O = Ay_XExl_FTIx(O, FTIxAy(J))
   Next
Src_XRmv_Mth = O
End Function

Function Src_XRmv_Ty(A$(), TyNm$) As String()
Src_XRmv_Ty = Ay_XExl_FTIx(A, DclTyNm_TyFTIx(A, TyNm))
End Function

Function Src_XRpl_Mth(A$(), MthNm$, NewMthLy$()) As String()
Dim OldMthLines$
   OldMthLines = SrcMthNm_MthLines(A, MthNm)
Dim NewMthLines$
   NewMthLines = JnCrLf(NewMthLy)
If OldMthLines = NewMthLines Then
   Src_XRpl_Mth = A
   Exit Function
End If
Dim O$()
   O = Src_XRmv_Mth(A, MthNm)
   PushAy O, NewMthLy
Src_XRpl_Mth = O

End Function

Property Get SrcTny() As String()
SrcTny = Db_SrcTny(CurrentDb)
End Property


Private Sub Z()
Z_Src_MthNmDic
Z_SrcMthNm_MthLines
Z_Src_MthDDNy
Z_Src_MthNy
MIde_Z_Src:
End Sub
