Attribute VB_Name = "MDao_Lnk_Import"
Option Compare Binary
Option Explicit

Private Property Get C_WhyBlkIsEr_MsgAy() As String()
Dim O$()
Push O, "The block is error because, it is none of these [RmkBlk SqBlk PrmBlk SwBlk]"
Push O, "SwBlk is all remark line or SwLin, which is started with ?"
Push O, "PrmBlk is all remark line or PrmLin, which is started with %"
Push O, "SqBlk is first non-remark begins with these [sel seldis drp upd] with optionally ?-Pffx"
Push O, "RmkBlk is all remark lines"
End Property

Function DSpecNm$(A)
DSpecNm = XTak_AftDotOrAll(Lin_T1(A))
End Function

Sub FnyIxAsg(Fny$(), FldLvs$, ParamArray OAp())
'FnyIxAsg=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = Ay_IxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub

Function FnySubIxAy(A$(), SubFny0) As Integer()
Dim SubFny$(): SubFny = CvNy(SubFny0)
If Sz(SubFny) = 0 Then Stop
Dim O%(), U&, J%
U = UB(SubFny)
ReSz O, U
For J = 0 To U
    O(J) = Ay_Ix(A, SubFny(J))
    If O(J) = -1 Then Stop
Next
FnySubIxAy = O
End Function

Sub FnyWhFldLvs(Fny$(), FldLvs$, ParamArray OAp())
'FnyWhFldLvs=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = Ay_IxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub

Private Sub HowToXEnsFirstTime_FmtSpec()
'No table-Spec
'No rec-Fmt
End Sub

Function ISpecINm$(A$)
ISpecINm = Lin_T1(A)
End Function

Private Function InActInpy(ASw() As LnkASw, FmSw() As LnkFmSw) As String()
Dim O$(), I, IFm As LnkFmSw, SwNm$, TF As Boolean
'If Sz(ASw) = 0 Then Exit Function

'For Each I In FmSw
'    Set IFm = I
'    SwNm = IFm.SwNm
'    TF = IFm.TF
'    If Not InActInpy__zSel(SwNm, TF, ASw) Then
'        PushAy O, Ssl_Sy(Rmv3T(IFm.Inpy))
'    End If
'Next
InActInpy = O
End Function

Private Function InActInpy__zSel(SwNm$, TF As Boolean, ASw() As LnkASw) As Boolean
'Dim IA As LnkASw, I
'For Each I In ASw
'    Set IA = I
'    If SwNm = IA.SwNm Then
'        InActInpy__zSel = IA.TF = TF
'        Exit Function
'    End If
'Next
'Stop
End Function

Sub IxAy_XAss_AllEleBet0ToU(IxAy, U)
Dim O(), Ix
For Each Ix In AyNz(IxAy)
    If 0 > Ix Or Ix > U Then
        XThw CSub, "IxAy has some ele not between 0 to U", "IxAy U", IxAy, U
    End If
Next
End Sub

Function IxAy_IsAllGE0(IxAy) As Boolean
Dim Ix
For Each Ix In AyNz(IxAy)
    If Ix < 0 Then Exit Function
Next
IxAy_IsAllGE0 = True
End Function

Function LyExt(NoT1$()) As String()
LyExt = LyXXX(NoT1, "Ext")
End Function

Function LyFld(NoT1$()) As String()
LyFld = LyXXX(NoT1, "Fld")
End Function

Private Function LyStuInp(NoT1$()) As String()
LyStuInp = LyXXX(NoT1, "StuInp")
End Function

Private Function LyXXX(NoT1$(), XXX$) As String()
LyXXX = Ay_XWh_RmvT1(NoT1, XXX)
End Function

Function MissingFnyChk(MissFny$(), ExistingFny$(), A As Database, T) As String()
If Sz(MissFny) = 0 Then Exit Function
Dim LnkVbl$
LnkVbl = Dbt_LnkVbl(A, T)
Const F1$ = "Excel file       : "
Const T1$ = "Worksheet        : "
Const C1$ = "Worksheet column : "
Const F2$ = "Database file: "
Const T2$ = "Table        : "
Const C2$ = "Field        : "
Dim X$, Y$, Z$, F0$, C0$, T0$
    AyAsg SplitVBar(LnkVbl), X, Y, Z
    Select Case X
    Case "LnkFb", "Lcl"
        F0 = F1
        T0 = T1
        C0 = C1
    Case "LnkFx"
        F0 = F2
        T0 = T2
        C0 = C2
    Case Else: Stop
    End Select
Dim O$()
    Dim I
    Push O, F0 & Y
    Push O, T0 & Z
    PushUnderLin O
    For Each I In ExistingFny
        Push O, C0 & XQuote_SqBkt(I)
    Next
    PushUnderLin O
    For Each I In MissFny
        Push O, C0 & XQuote_SqBkt(I)
    Next
'    PushMsgUnderLinDbl O, QQ_Fmt("Above ? are missing", D)
Stop '
MissingFnyChk = O
End Function

Sub XMap_NDrive()
XRmv_NDrive
Shell "Subst N: c:\users\user\desktop\MHD"
End Sub

Sub XRmv_NDrive()
Shell "Subst /d N:"
End Sub

Function NewIntSeq(N&, Optional IsFmOne As Boolean) As Integer()
Dim O%(): ReDim O(N - 1)
Dim J&
If IsFmOne Then
    For J = 0 To N - 1
        O(J) = J + 1
    Next
Else
    For J = 0 To N - 1
        O(J) = J
    Next
End If
NewIntSeq = O
End Function

Function NewIpFil(Ly$()) As LnkIpFil()
If Sz(Ly) = 0 Then Exit Function
Dim O() As LnkIpFil, J%, L, Ay
ReDim O(UB(Ly))
For Each L In Ly
'    Ay = AyT1Rst(Ssl_Sy(L))
    Stop '
'    Set O(J) = NewLnkIpFil(L)
    O(J).Fil = Ay(0)
    'O(J).Inpy = CvSy(Ay(1))
    J = J + 1
Next
NewIpFil = O
End Function

Function New_SimTy(SimTyStr$) As eSimTy
Dim O As eSimTy
Select Case UCase(SimTyStr)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
New_SimTy = O
End Function

Function New_StExt(Lin) As LnkStExt
'Dim O As New LnkStExt
'With O
'    AyAsg Lin3TAy(Lin), .LikInp, .F, , .Ext
'End With
'Set New_StExt = O
End Function

Private Function New_StFld(Lin) As LnkStFld
'Dim O As New LnkStFld, A$
'With O
'    AyAsg Lin2TAy(Lin), .Stu, , A
'    .Fny = Ssl_Sy(A)
'End With
'Set New_StFld = O
End Function

Function U_Sy(U&) As String()
Dim O$()
If U > 0 Then ReDim O(U)
U_Sy = O
End Function

Function Nm_NxtSeqNm$(A, Optional NDig% = 3) _
'Nm-A can be XXX or XXX_nn
'   If XXX, return XXX_001   '<-- # of zero depends on NDig
'   If XXX_nn, return XXX_mm '<-- mm is nn+1, # of digit of nn and mm depends on NDig
If NDig = 0 Then Stop
Dim R$
    R = Right(A, NDig + 1)

If Left(R, 1) <> "_" Then GoTo Case1
If Not IsNumeric(Mid(R, 2)) Then GoTo Case1

Dim L$: L = Left(A, Len(A) - NDig)
Dim Nxt%: Nxt = Val(Mid(R, 2)) + 1
Nm_NxtSeqNm = Left(A, Len(A) - NDig) + XPad0(Nxt, NDig)
Exit Function

Case1:
    Nm_NxtSeqNm = A & "_" & StrDup(NDig - 1, "0") & "1"
End Function

Function Ny0SqBktCsv$(A)
Dim B$(), C$()
B = CvNy(A)
C = Ay_XQuote_SqBkt(B)
Ny0SqBktCsv = JnComma(C)
End Function

Private Function NyEy(Ny$(), A() As LnkStEle) As String()

End Function

Function NyEySqy$(Ny$(), Ey$())
Ay_IsEqSz_XAss Ny, Ey
If Ay_IsEq(Ny, Ey) Then
    NyEySqy = "Select" & vbCrLf & "    " & JnComma(Ny)
    Exit Function
End If
Dim N$()
    N = Ay_XAlign_L(Ny)
Dim E$()
    Dim J%
    E = Ey
    For J = 0 To UB(E)
        If E(J) <> "" Then E(J) = XQuote_SqBkt(E(J))
    Next
    E = Ay_XAlign_L(E)
    For J = 0 To UB(E)
        If Trim(E(J)) <> "" Then E(J) = E(J) & " As "
    Next
    E = Ay_XAlign_L(E)
Dim O$()
    O = AyAB_XJn(E, N)
NyEySqy = Join(O, "," & vbCrLf)
End Function

Function NyLnxAy(Ny0) As String()
'It is to return 2 lines with
'first line is 0   1     2 ..., where 0,1,2.. are ix of A$()
'second line is each element of A$() separated by A
'Eg, A$() = "A BBBB CCC DD"
'return 2 lines of
'0 1    2   3
'A BBBB CCC DD
Dim Ny$(): Ny = CvNy(Ny0)
If Sz(Ny) = 0 Then Exit Function
Dim A1$()
Dim A2$()
Dim U&: U = UB(Ny)
ReSz A1, U
ReSz A2, U
Dim O$(), J%, L$, W%
For J = 0 To U
    L = Len(Ny(J))
    W = Max(L, Len(J))
    A1(J) = XAlignL(J, W)
    A2(J) = XAlignL(Ny(J), W)
Next
Push O, JnSpc(A1)
Push O, JnSpc(A2)
NyLnxAy = O
End Function

Function PfxSsl_Sy(A) As String()
Dim Ay$(), Pfx$
Ay = Ssl_Sy(A)
Pfx = Ay_XShf_(Ay)
PfxSsl_Sy = Ay_XAdd_Pfx(Ay, Pfx)
End Function

Function Prepend$(S, P)
If S <> "" Then Prepend = P & S
End Function

Function PrependDash$(S)
PrependDash = Prepend(S, "-")
End Function

Function ReSeqSpec_Fny(A) As String()
Dim Ay$(), D As Dictionary, O$(), L1$, L
Ay = SplitVBar(A)
If Sz(Ay) = 0 Then Exit Function
L1 = Ay_XShf_(Ay)
Set D = LyTRst_Dic(Ay)
For Each L In Ssl_Sy(L1)
    If D.Exists(L) Then
        Push O, D(L)
    Else
        Push O, L
    End If
Next
ReSeqSpec_Fny = Ssl_Sy(JnSpc(O))
End Function

Function ReSeqSpec_OLinFldAy(A) As String()
Dim B$()
B = SplitVBar(A)
Ay_XShf_ B
ReSeqSpec_OLinFldAy = Ay_XTak_T1(B)
End Function

Function ReSeqSpec_OutLin(A, F) As Byte
Dim Ay$(), Ssl, J%
Ay = SplitVBar(A)
If SslHas(Ay(0), F) Then Exit Function
For J = 1 To UB(Ay)
    Select Case SslIx(Ssl, F)
    Case 0: Stop
    Case Is > 0
        ReSeqSpec_OutLin = 2
    End Select
Next
End Function

Sub ResClr(A$)
DbResClr CurDb, A
End Sub

Private Function SelIntoAy(ActInpy$(), A As LnkSpec) As String()
Dim Inp$, I, J%, O()
ReDim O(UB(ActInpy))
For Each I In ActInpy
'    Set O(J) = New SqlSelInto
    With O(J)
        Inp = I
'        .Ny = InpNy(Inp, A.StInp, A.StFld)
'        .Ey = NyEy(.Ny, A.StEle)
'        .Fm = ">" & Inp
'        .Into = "#I" & Inp
'        .Wh = InpWhBExpr(Inp, A.FmWh)
    End With
    J = J + 1
Next
'SelIntoAy = O
End Function

Sub SelRg_SetXorEmpty(A As Range)
Dim I
For Each I In A
    
Next
End Sub

Function SetSqpFmt$(Fny$(), Vy())
Dim A$: GoSub X_A
SetSqpFmt = vbCrLf & "  Set" & vbCrLf & A
Exit Function
X_A:
    Dim L$(): L = FnyAlignQuote(Fny)
    Dim R$(): GoSub X_R
    Dim J%, O$(), S$
    S = Space(4)
    For J = 0 To UB(L)
        Push O, S & L(J) & "= " & R(J)
    Next
    A = JnCrLf(O)
    Return
X_R:
    R = Ay_XAlign_L(Vy_XQuote_Sql(Vy))
    Dim J1%
    For J1 = 0 To UB(R) - 1
        R(J1) = R(J1) + ","
    Next
    Return
End Function

Function T0F2LinHasTF(A, T$, F$) As Boolean
Dim TLik$, FLikSsl$
Lin_TRstAsg A, TLik, FLikSsl
If T Like TLik Then
    If XLik_Likss(F, FLikSsl) Then
        T0F2LinHasTF = True
        Exit Function
    End If
End If
End Function

Function TFLinHasPk(A$) As Boolean
TFLinHasPk = HasSubStr(A, " * ")
End Function

Function TFLinHasSk(A$) As Boolean
TFLinHasSk = HasSubStr(A, " | ")
End Function

Function TFTyChkMsg$(T, F, Ty As DAO.DataTypeEnum, ExpTyAy() As DAO.DataTypeEnum)
'Dbtf_TyMsg = QQ_Fmt("Table[?] field[?] has type[?].  It should be type[?].", T, F, S1, S2)

End Function

Function TimSz_XTSz$(A As Date, Sz&)
TimSz_XTSz = Dte_DTim(A) & "." & Sz
End Function

Function TkIsExist(T, K&) As Boolean
TkIsExist = Dbt_Existk(CurDb, T, K)
End Function

Property Get UniqFny() As String()
Stop '
'Dim I, M As LABC, O$()
'If IsEmp Then Exit Property
'For Each I In A
'    Set M = I
'    PushNoDupAy O, M.Fny
'Next
'UniqFny = O
End Property

Function XFLinX$(A, F)
Dim X$, FLikss$
Lin_TRstAsg A, X, FLikss
If XLik_Likss(F, FLikss) Then XFLinX = X
End Function

Function XFyX$(A$(), F)
Dim L
For Each L In AyNz(A)
    XFyX = XFLinX(L, F)
    If XFyX <> "" Then Exit Function
Next
End Function

Private Property Get ZZCrdTyLvs$()
ZZCrdTyLvs = "1 2 3"
End Property

Private Sub ZZ_Ap_DtAy()
Dim A() As Dt
A = Ap_DtAy(Samp_Dt1, Samp_Dt2)
Stop
End Sub

Private Sub ZZ_ErAyzFx_WsMissingCol()
'" [Material]             As Sku," & _
'" [Plant]                As Whs," & _
'" [Storage Location]     As Loc," & _
'" [Batch]                As BchNo," & _
'" [Unrestricted]         As OH " & _

End Sub

Private Sub ZZ_ReSeqSpec_Fny()
Ay_XBrw ReSeqSpec_Fny("Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
" Flg IsAlert IsWithSku |" & _
" Key Sku PstMth PstDte |" & _
" Bch BchNo BchPermitDate BchPermit |" & _
" Las LasBchNo LasPermitDate LasPermit |" & _
" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" Uom Des StkUom Ac_U")
End Sub

Private Sub ZZ_SslSqBktCsv()
Debug.Print SslSqBktCsv("a b c")
End Sub

Private Sub ZZ_TF_Des()
TF_Des("Att", "AttNm") = "AttNm"
End Sub

Private Sub Z_Ap_JnSemiColon()
Act = Ap_JnSemiColon(" ", "")
Ept = ""
C
End Sub

Private Sub Z_Db_XAdd_Td()
Dim A As DAO.TableDef
Tbl_XDrp "Tmp"
Set A = Db_XAdd_Td(CurDb, TmpTd)
Tbl_XDrp "Tmp"
End Sub

Private Sub Z_Dbtf_XAdd_Expr()
Tbl_XDrp "Tmp"
Dim A As DAO.TableDef
Set A = Db_XAdd_Td(CurDb, TmpTd)
Dbtf_XAdd_Expr CurDb, "Tmp", "F2", "[F1]+"" hello!"""
Tbl_XDrp "Tmp"
End Sub

Private Sub Z_Drs_XEns_RplDbt()
Dim Db As Database, D1 As Drs, D2 As Drs
Set Db = TmpDb
Set D1 = Samp_Drs
Drs_XEns_RplDbt D1, Db, "T"
Set D2 = Dbt_Drs(Db, "T")
Ass Ay_IsEq(D1, D2)
Db_XKill Db
End Sub

Private Sub Z_Drs_XInsUpd_Dbt()
Dim Db As Database, T$, A As Drs, TFb$
    TFb = TmpFb("Tst", "Drs_XInsUpd_Dbt")
    Set Db = Fb_XCrt(TFb)
T = "Tmp"
Db.Execute "Create Table Tmp (A Int, B Int, C Int)"
Db.Execute TblSkFF_CrtSkSql("Tmp", "A")
'DbSqy_XRun Db, InsDrApSqy("Tmp", "A B C", Array(1, 3, 4), Array(3, 4, 5))
Set A = New_Drs("A B C", CvAy(Array(Array(1, 2, 3), Array(2, 3, 4))))

Ept = Array(Array(1&, 2&, 3&), Array(2&, 3&, 4&), Array(3&, 4&, 5&))
GoSub Tst
Db.Close
Kill TFb
Exit Sub
Tst:
    Drs_XInsUpd_Dbt A, Db, T
    Act = Dbt_Dry(Db, T)
    C
    Return
End Sub

Private Sub Z_LblSeqAy()
Dim Act$(), A, N%, Exp$()
A = "Lbl"
N = 10
Exp = Ssl_Sy("Lbl01 Lbl02 Lbl03 Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10")
Act = LblSeqAy(A, N)
Ass Ay_IsEq(Act, Exp)
End Sub

Private Sub Z_PfxSsl_Sy()
Dim A$, Exp$()
A = "A B C D"
Exp = Ssl_Sy("AB AC AD")
GoSub Tst
Exit Sub
Tst:
Dim Act$()
Act = PfxSsl_Sy(A)
Debug.Assert Ay_IsEq(Act, Exp)
Return
End Sub

Private Sub Z_SclChk()
Dim A$, Ny0
A = "Req;Alw;Sz=1"
Ny0 = VdtEleSclNmSsl
Ept = EmpSy
Push Ept, "There are [invalid-SclNy] in given [scl] under these [valid-SclNy]."
Push Ept, "    [invalid-SclNy] : Alw"
Push Ept, "    [scl]           : Req;Alw;Sz=1"
Push Ept, "    [valid-SclNy]   : Req AlwZLen Sz Dft VRul VTxt Des Expr"
GoSub Tst
Exit Sub
Tst:
    Act = SclChk(A, Ny0)
    C
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$()
Dim C$
Dim D()
Dim E As Database
Dim F&
Dim G As Boolean
Dim H%
Dim I As Range
Dim J As DAO.DataTypeEnum
Dim K() As DAO.DataTypeEnum
Dim L As Date

DSpecNm A
FnyIxAsg B, C, D
FnySubIxAy B, A
FnyWhFldLvs B, C, D
ISpecINm C
IxAy_IsAllGE0 A
LyExt B
LyFld B
MissingFnyChk B, B, E, A
XMap_NDrive
XRmv_NDrive
NewIntSeq F, G
NewIpFil B
New_SimTy C
New_StExt A
Nm_NxtSeqNm A, H
Ny0SqBktCsv A
NyEySqy B, B
NyLnxAy A
PfxSsl_Sy A
Prepend A, A
PrependDash A
ReSeqSpec_Fny A
ReSeqSpec_OLinFldAy A
ReSeqSpec_OutLin A, A
ResClr C
SelRg_SetXorEmpty I
SetSqpFmt B, D
T0F2LinHasTF A, C, C
TFLinHasPk C
TFLinHasSk C
TFTyChkMsg A, A, J, K
TimSz_XTSz L, F
TkIsExist A, F
XFLinX A, A
XFyX B, A
QQInBExpr_F_Ay A, C, G
End Sub

Private Sub Z()
Z_Ap_JnSemiColon
Z_Db_XAdd_Td
Z_Dbtf_XAdd_Expr
Z_Drs_XEns_RplDbt
Z_Drs_XInsUpd_Dbt
Z_LblSeqAy
Z_PfxSsl_Sy
Z_SclChk
End Sub
