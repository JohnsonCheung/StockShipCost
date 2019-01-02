Attribute VB_Name = "MTp_Sq_Sq"
Option Compare Binary
Option Explicit
Private Enum eStmtTy
    eUpdStmt = 1
    eDrpStmt = 2
    eSelStmt = 3
End Enum
Const U_Into$ = "INTO"
Const U_Sel$ = "SEL"
Const U_SelDis$ = "SELECT DISTINCT"
Const U_Fm$ = "FM"
Const U_Gp$ = "GP"
Const U_Wh$ = "WH"
Const U_And$ = "AND"
Const U_Jn$ = "JN"
Const U_LeftJn$ = "LEFT JOIN"
Private Pm As Dictionary
Private StmtSw As Dictionary
Private FldSw As Dictionary

Private Sub AAMain()
Z_SqBlkAy_SqyRslt
End Sub

Private Function Sq_Ly_SqlRslt_DRP(A$()) As SqlRslt
End Function

Private Function Sq_Ly_SqlRslt_SEL(A$(), ExprDic As Dictionary) As SqlRslt
Dim O$()
Dim Er$()
    Dim E As Dictionary
    Set E = ExprDic
    Dim I, J%, B$(), L$
    B = AyReverseI(A)
    PushI O, XSel(Pop(B), E)
    PushI O, Into(RmvT1(Pop(B)))
'    PushI O, QQFm_T(RmvT1(Pop(B)))
    PushIAy O, XJnOrLeftJn(PopJnOrLeftJn(B), E)
    L = PopWh(B)
    If L <> "" Then
        PushI O, XWh(L, E)
        PushIAy O, XAnd(PopAnd(B), E)
    End If
    PushI O, XGp(PopGp(B), E)
Dim OO As SqlRslt
    OO.Sql = JnCrLf(O)
    OO.Er = Er
Sq_Ly_SqlRslt_SEL = OO
End Function
Private Function Sq_Ly_XRmv_ExprLin(A) As String()

End Function

Private Function Sq_Ly_SqlRslt(A) As SqlRslt
Dim Ty As eStmtTy
    Ty = Sq_Ly_StmtTy(A)
If Sq_Ly_IsSkip(A, Ty) Then Exit Function
Dim Sq_LyNoExprLin$(), ExprDic As Dictionary
Set ExprDic = Sq_Ly_ExprDic(A)
Sq_LyNoExprLin = Sq_Ly_XRmv_ExprLin(A)
Dim O As SqlRslt
    Select Case Ty
    Case eUpdStmt: O = Sq_Ly_SqlRslt_UPD(Sq_LyNoExprLin, ExprDic)
    Case eDrpStmt: O = Sq_Ly_SqlRslt_DRP(Sq_LyNoExprLin)
    Case eSelStmt: O = Sq_Ly_SqlRslt_SEL(Sq_LyNoExprLin, ExprDic)
    Case Else: Stop
    End Select
Sq_Ly_SqlRslt = O
End Function

Private Function Sq_Ly_SqlRslt_UPD(A$(), E As Dictionary) As SqlRslt

End Function

Private Function Fny_XWh_Active(A$()) As String()
Dim F
For Each F In A
    If FldSw.Exists(F) Then PushI Fny_XWh_Active, F
Next
End Function

Private Function Sq_Ly_ExprDic(A) As Dictionary
Dim Expr$(), M As AyPair
M = Ay_XBrk_BY_ELE(A, "$")
Set Sq_Ly_ExprDic = New_Dic_LY(CvSy(M.B))
End Function

Private Function Sq_Ly_ExprDicAy(Fny$(), E As Dictionary) As String()
Dim F, M$
For Each F In Fny
    If E.Exists(F) Then
        M = E(F)
    Else
        M = F
    End If
    PushI Sq_Ly_ExprDicAy, M
Next
End Function

Private Function Sq_Ly_IsSkip(A, Ty As eStmtTy) As Boolean
Sq_Ly_IsSkip = StmtSw.Exists(Sq_Ly_StmtSwKey(A, Ty))
End Function

Private Function Sq_Ly_StmtTy(A) As eStmtTy
Dim L$
L = UCase(XRmv_Pfx(Lin_T1(A(0)), "?"))
Select Case L
Case "SEL": Sq_Ly_StmtTy = eSelStmt
Case "UPD": Sq_Ly_StmtTy = eUpdStmt
Case "DRP": Sq_Ly_StmtTy = eDrpStmt
Case Else: Stop
End Select
End Function

Private Function Sq_Ly_StmtSwKey$(A, Ty As eStmtTy)
Select Case Ty
Case eStmtTy.eSelStmt: Sq_Ly_StmtSwKey = Sq_Ly_StmtSwKey_SEL(A)
Case eStmtTy.eUpdStmt: Sq_Ly_StmtSwKey = Sq_Ly_StmtSwKey_UPD(A)
Case Else: Stop
End Select
End Function

Private Function Sq_Ly_StmtSwKey_SEL$(A)
Sq_Ly_StmtSwKey_SEL = Ay_FstRmvT1(A, "FM")
End Function

Private Function Sq_Ly_StmtSwKey_UPD$(A)
Dim Lin1$
    Lin1 = A(0)
If XRmv_Pfx(XShf_Term(Lin1), "?") <> "upd" Then Stop
Sq_Ly_StmtSwKey_UPD = Lin1
End Function

Private Function FndVy(K, E As Dictionary, OVy$(), OQ$)
'Return true if not found
End Function

Private Function FndValPair(K, E As Dictionary, OV1, OV2)
'Return true if not found
End Function

Private Function IsXXX(A$(), XXX$) As Boolean
IsXXX = UCase(Lin_T1(A(UB(A)))) = XXX
End Function


Private Function MsgAp_Lin_TyEr(A As Lnx) As String()


End Function

Private Function MsgMustBeIntoLin$(A As Lnx)

End Function

Private Function MsgMustBeSelorSelDis$(A As Lnx)

End Function

Private Function MsgMustNotHasSpcInTbl_NmOfIntoLin$(A As Lnx)

End Function


Private Property Get SampExprDic() As Dictionary
Dim O$()
PushI O, "A XX"
PushI O, "B BB"
PushI O, "C DD"
PushI O, "E FF"
Set SampExprDic = New_Dic_LY(O)
End Property

Private Property Get SampSqLnxAy() As Lnx()
Dim O$()
PushI O, "sel ?MbrCnt RecCnt TxCnt Qty Amt"
PushI O, "into #Cnt"
PushI O, "fm   #Tx"
PushI O, "wh   RecCnt bet @XX @XX"
PushI O, "and  RecCnt bet @XX @XX"

PushI O, "$"
PushI O, "?MbrCnt ?Count(Distinct Mbr)"
PushI O, "RecCnt  Count(*)"
PushI O, "TxCnt   Sum(TxCnt)"
PushI O, "Qty     Sum(Qty)"
PushI O, "Amt     Sum(Amt)"
SampSqLnxAy = Ly_LnxAy(O)
End Property
Private Function GpAy_Er(A() As Gp, OLyAy()) As String()

End Function

Function SqBlkAy_SqyRslt(SqGpAy() As Gp, PmDic As Dictionary, StmtSwDic As Dictionary, FldSwDic As Dictionary) As SqyRslt
Set Pm = Pm
Set StmtSw = StmtSwDic
Set FldSw = FldSwDic
Dim LyRslt As LyRslt
    LyRslt = SqGpAy_LyRslt(SqGpAy)
Dim Ly, O As SqyRslt
For Each Ly In AyNz(LyRslt.Ly)
    O = SqyRslt_XAdd_SqlRslt(O, Sq_Ly_SqlRslt(CvSy(Ly)))
Next
End Function
Private Function SqGpAy_LyRslt(A() As Gp) As LyRslt

End Function

Function SqyRslt_XAdd_SqlRslt(A As SqyRslt, B As SqlRslt) As SqyRslt
Dim O As SqyRslt
O = A

End Function
Function MsgAndLinOp_ShouldBe_BetOrIn(A)

End Function
Private Function XAnd(A$(), E As Dictionary)
'and f bet xx xx
'and f in xx
Dim F$, I, L$, Ix%, M As Lnx
For Each I In AyNz(A)
    Set M = I
    LnxAsg M, L, Ix
    If XShf_Term(L) <> "and" Then Stop
    F = XShf_Term(L)
    Select Case XShf_Term(L)
    Case "bet":
    Case "in"
    Case Else: Stop
    End Select
Next
End Function


Private Function XGp$(L$, E As Dictionary)
If L = "" Then Exit Function
Dim ExprAy$(), Ay$()
Stop
'    ExprAy = DicSelIntoSy(EDic, Ay)
'XGp = SqpGp(ExprAy)
End Function

Private Function XJnOrLeftJn(A$(), E As Dictionary) As String()

End Function

Private Function PopJnOrLeftJn(A$()) As String()
PopJnOrLeftJn = PopMulXorYOpt(A, U_Jn, U_LeftJn)
End Function

Private Function PopXXXOpt$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else return ''
If Sz(A) = 0 Then Exit Function
PopXXXOpt = PopXXX(A, XXX)
End Function

Private Function PopXXX$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else Stop
Dim L$: L = A(UB(A))
If XRmv_Pfx(Lin_T1(L), "?") = XXX Then
    PopXXX = RmvT1(L)
    Pop A
End If
End Function

Private Function PopGp$(A$())
PopGp = PopXXXOpt(A, U_Gp)
End Function

Private Function PopWh$(A$())
PopWh = PopXXXOpt(A, U_Wh)
End Function

Private Function PopAnd(A$()) As String()
PopAnd = PopMulXXX(A, U_And)
End Function

Private Function PopXorYOpt$(A$(), X$, Y$)
Dim L$
L = PopXXXOpt(A, X): If L <> "" Then PopXorYOpt = L: Exit Function
PopXorYOpt = PopXXXOpt(A, Y)
End Function

Private Function PopMulXorYOpt(A$(), X$, Y$) As String()
Dim J%, L$
While Sz(A) > 0
    J = J + 1: If J > 1000 Then Stop
    L = PopXorYOpt(A, X, Y)
    If L = "" Then Exit Function
    PushI PopMulXorYOpt, L
Wend
End Function

Private Function PopMulXXX(A$(), XXX$) As String()
Dim J%
While Sz(A) > 0
    J = J + 1: If J > 1000 Then Stop
    If Not IsXXX(A, XXX) Then Exit Function
    PushObj PopMulXXX, Pop(A)
Wend
End Function

Private Function XSel$(A$, E As Dictionary)
Dim Fny$()
    Dim T1$, L$
    L = A
    T1 = XRmv_Pfx(XShf_Term(L), "?")
    Fny = XSelFny(Ssl_Sy(L), FldSw)
Select Case T1
'Case U_Sel:    XSel = X.Sel_Fny_EDic(Fny, E)
'Case U_SelDis: XSel = X.Sel_Fny_EDic(Fny, E, IsDis:=True)
Case Else: Stop
End Select
End Function
Private Function XSelFny(Fny$(), FldSw As Dictionary) As String()
Dim F
For Each F In Fny
    If XTak_FstChr(F) = "?" Then
        If Not FldSw.Exists(F) Then Stop
        If FldSw(F) Then PushI XSelFny, F
    Else
        PushI XSelFny, F
    End If
Next
End Function

Private Function XSet(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XUpd(A As Lnx, E As Dictionary, OEr$())

End Function
Private Function XWh$(L$, E As Dictionary)
'L is following
'  ?Fld in @ValLis  -
'  ?Fld bet @V1 @V2
Dim F$, Vy$(), V1, V2, IsBet As Boolean
If IsBet Then
    If Not FndValPair(F, E, V1, V2) Then Exit Function
    'XWh = QQWh_BExprBet(F, V1, V2)
    Exit Function
End If
'If Not FndVy(F, E, Vy, Q) Then Exit Function
'XWh = QQWh_BExprFldInAy(F, Vy)
End Function

Private Function XWhBetNbr$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XWhExpr(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XWhInNbrLis$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Sub Z_Sq_Ly_SqlRslt_SEL()
Dim E As Dictionary, Ly$(), Act As SqlRslt
'---
Erase Ly
    Push Ly, "?XX Fld-XX"
    Push Ly, "BB Fld-BB-LINE-1"
    Push Ly, "BB Fld-BB-LINE-2"
    Set E = New_Dic_LY(Ly)           '<== Set ExprDic
Erase Ly
    Set FldSw = New Dictionary
    FldSw.Add "?XX", False       '<=== Set FldSw
Erase Ly
    Erase Ly
    PushI Ly, "sel ?XX BB CC"
    PushI Ly, "into #AA"
    PushI Ly, "fm   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "wh   A bet $a $b"
    PushI Ly, "and  B in $c"
    PushI Ly, "gp   D C"        '<== Sq_Ly
GoSub Tst
Exit Sub
Tst:
    Act = Sq_Ly_SqlRslt_SEL(Ly, E)
    C
    Return
End Sub

Private Sub Z_Sq_Ly_ExprDic()
Dim Ly$()
Dim D As New Dictionary
'-----

Erase Ly
PushI Ly, "aaa bbb"
PushI Ly, "111 222"
PushI Ly, "$"
PushI Ly, "A B0"
PushI Ly, "A B1"
PushI Ly, "A B2"
PushI Ly, "B B0"
D.RemoveAll
    D.Add "A", JnCrLf(Ssl_Sy("B0 B1 B2"))
    D.Add "B", "B0"
    Set Ept = D
GoSub Tst
Exit Sub
Tst:
    Set Act = Sq_Ly_ExprDic(Ly)
    Ass Dic_IsEq(CvDic(Act), CvDic(Ept))
    
    Return
End Sub

Private Sub Z_SqBlkAy_SqyRslt()
Dim A() As Gp, Pm As Dictionary, StmtSw As Dictionary, FldSw As Dictionary, Er$()
GoSub Dta1
GoSub Tst
Return
Tst:
    Dim Act As SqyRslt
    Act = SqBlkAy_SqyRslt(A, Pm, StmtSw, FldSw)
    C
    Return
Dta1:
    Return
End Sub

Private Sub Z_XSel()
Dim A$, E As Dictionary
A = "dsklfj"
Set E = SampExprDic
GoSub Tst
Exit Sub
Tst:
    Act = XSel(A, E)
    C
    Return
End Sub

Private Sub Z_Sq_Ly_StmtSwKey()
Dim Ly$(), Ty As eStmtTy
'---
PushI Ly, "sel sdflk"
PushI Ly, "fm AA BB"
Ept = "AA BB"
Ty = eSelStmt
GoSub Tst
'---
Erase Ly
PushI Ly, "?upd XX BB"
PushI Ly, "fm dsklf dsfl"
Ept = "XX BB"
Ty = eUpdStmt
GoSub Tst
Exit Sub
Tst:
    Act = Sq_Ly_StmtSwKey(Ly, Ty)
    C
    Return
End Sub


Private Sub Z()
Z_Sq_Ly_SqlRslt_SEL
Z_Sq_Ly_ExprDic
Z_SqBlkAy_SqyRslt
Z_Sq_Ly_StmtSwKey
Z_XSel
MTp_Sq_Sq:
End Sub
