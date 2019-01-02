Attribute VB_Name = "MXls_Z_Lo_Fmt"
Option Compare Binary
Option Explicit
Public Const M_Val_IsNonNum$ = "Lx(?) has Val(?) should be a number"
Public Const M_Val_IsNonLng$ = "Lx(?) has Val(?) should be a 'Long' number"
Public Const M_Val_ShouldBet$ = "Lx(?) has Val(?) should be between [?] and [?]"
Public Const M_Fld_IsInValid$ = "Lx(?) Fld(?) is invalid.  Not found in Fny"
Public Const M_Fld_IsDup$ = "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"
Public Const M_Nm_LinHasNoVal$ = "Lx(?) is Nm-Lin, it has no value"
Public Const M_Nm_NoNmLin$ = "Nm-Lin is missing"
Public Const M_Nm_ExcessLin$ = "LX(?) is excess due to Nm-Lin is found above"
Public Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Public Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Public Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"

Const M_Fny$ = "Lin_Ty(?) has these Fld(?) in not Fny"
Const M_Bdr_ExcessFld$ = "These Fld(?) in [Bdr ?] already exists in [Bdr ?], they are skipped in setting border"
Const M_Bdr_ExcessLin$ = "These Fld(?) in [Bdr ?] already exists in [Bdr ?], they are skipped in setting border"
Const M_CorVal$ = "In Lin(?)-Color(?), color cannot convert to long"
Const M_Fld_IsAvg_FndInSum$ = "Lin(?)-Fld(?), which is TAvg-Fld, but also found in TSum-Lx(?)"
Const M_Fld_IsCnt_FndInSum$ = "Lin(?)-Fld(?), which is TCnt-Fld, but also found in TSum-Lx(?)"
Const M_Fld_IsCnt_FndInAvg$ = "Lin(?)-Fld(?), which is TCnt-Fld, but also found in TAvg-Lx(?)"
Const M_Bet_Should2Term = "Lin(?)-Fld(?) is Bet-Line.  It should have 2 terms"
Const M_Bet_InvalidTerm = "Lin(?)-Fld(?) is Bet-Line.  It has invalid term(?)"
Const M_Dup$ = "Lin(?)-Fld(?) is duplicated.  The line is skipped"
Dim A_Lo As ListObject, A_Fny$(), A_Lo_Fmtr$()
Dim Align$(), Bdr$(), Tot, Bet$()
Dim Wdt$(), Fmt$(), Lvl$(), Cor$()
Dim Tit$(), Fml$(), Lbl$()

Private Sub AAMain()
Z_Lo_Fmt
End Sub

Private Sub XSet_Fmt1(L)
Dim F
Dim Fmt$, Fldss$
For Each F In A_Fny
    Lin_TRstAsg L, Fmt, Fldss
    If XLik_Likss(F, Fldss) Then
'        Lc_XSet_Align A_Lo, F, Fmt
        Exit Sub
    End If
Next
End Sub

Private Sub XSet_Fmt()
Dim L
For Each L In AyNz(Fmt)
    XSet_Fmt1 L
Next
End Sub

Private Property Get WEr_Align() As String()
WEr_Align = Ap_Sy(WEr_AlignLin, WEr_AlignFny)
End Property

Private Property Get WEr_AlignFny() As String()
Dim WEr_Fny$()
    Dim AlignFny$()
    AlignFny = Ay_XWh_Dist(SSAy_Sy(Ay_XRmv_TT(Align)))
    WEr_Fny = AyMinus(AlignFny, A_Fny)
End Property

Private Property Get WEr_AlignLin() As String()
WEr_AlignLin = WMsg_AlignLin(Ay_XExl_T1Ay(Align, "Left Right Center"))
End Property

Private Function WEr_Bdr1(X$) As String()
'Return FldAy from Bdr & X
Dim FldssAy$(): FldssAy = SSAy_Sy(Ay_XWh_RmvT1(Bdr, X))
End Function

Private Property Get WEr_Bdr() As String()
WEr_Bdr = Ap_Sy(WEr_BdrExcessFld, WEr_BdrExcessLin, WEr_BdrDup, WEr_BdrFld)
End Property

Private Property Get WEr_BdrDup() As String()
WEr_BdrDup = WMsg_Dup(AyDupT1(Bdr), Bdr)
End Property

Private Property Get WEr_BdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = WEr_Bdr1("Left")
RfNy = WEr_Bdr1("Right")
CFny = WEr_Bdr1("Center")
PushIAy WEr_BdrExcessFld, QQ_Fmt(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
PushIAy WEr_BdrExcessFld, QQ_Fmt(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
PushIAy WEr_BdrExcessFld, QQ_Fmt(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Property

Private Property Get WEr_BdrExcessLin() As String()
Dim L
For Each L In AyNz(Ay_XExl_T1Ay(Bdr, "Left Right Center"))
    PushI WEr_BdrExcessLin, QQ_Fmt(M_Bdr_ExcessLin, L)
Next
End Property

Private Property Get WEr_BdrFld() As String()
Dim Fny$(): Fny = Ap_Sy(WEr_Bdr1("Left"), WEr_Bdr1("Right"), WEr_Bdr1("Center"))
WEr_BdrFld = WMsg_Fny(Fny, "Bdr")
End Property

Private Property Get WEr_Bet() As String()
WEr_Bet = Ap_Sy(WEr_BetDup, WEr_BetFny, WEr_BetTermCnt)
End Property

Private Property Get WEr_BetDup() As String()
WEr_BetDup = WMsg_Dup(AyDupT1(Bet), Bet)
End Property

Private Property Get WEr_BetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return WEr_ of M_Bet_* if any
End Property

Private Property Get WEr_BetTermCnt() As String()
Dim L
For Each L In AyNz(Bet)
    If Sz(Ssl_Sy(L)) <> 3 Then
        PushI WEr_BetTermCnt, WMsg_BetTermCnt(L, 3)
    End If
Next
End Property

Private Property Get WEr_Cor() As String()
Dim L$()
L = Cor
WEr_Cor = Ap_Sy(WEr_CorDup(L), WEr_CorFld(L), WEr_CorVal(L))
Cor = L
End Property

Private Function WEr_CorDup(IO$()) As String()

End Function

Private Function WEr_CorFld(IO$()) As String()

End Function

Private Function WEr_CorVal1$(L)
Dim Cor$
Cor = Lin_T1(L)
If IsEmpty(CvColr(L)) Then
    WEr_CorVal1 = QQ_Fmt(M_CorVal, L, Cor)
End If
If CanCvLng(Cor) Then Exit Function
End Function

Private Function WEr_CorVal(IO$()) As String()
Dim Msg$(), WEr_$(), L
For Each L In IO
    PushI Msg, WEr_CorVal1(L)
Next
IO = Ay_XWh_NoEr(IO, Msg, WEr_)
End Function

Private Property Get WEr_Fml() As String()
WEr_Fml = Ap_Sy(WEr_FmlDup, WEr_FmlFny)
End Property

Private Property Get WEr_FmlDup() As String()
WEr_FmlDup = WMsg_Dup(AyDupT1(Fml), Fml)
End Property

Private Property Get WEr_FmlFny() As String()
'WEr_FmlFny = AyMinus(FmlNy(Fml), A_Fny)
End Property

Private Property Get WEr_Fmt() As String()

End Property

Private Property Get WEr_Lbl() As String()

End Property

Private Property Get WEr_Lvl() As String()

End Property

Private Property Get WEr_Tit() As String()

End Property

Private Property Get WEr_Tot() As String()
Dim L
For Each L In AyNz(Tot)
'    A = Avg(J)
'    Ix = Ay_Ix(Sum, A)
'    If Ix >= 0 Then
'        Msg = QQ_Fmt(M_Fld_IsAvg_FndInSum, AvgLxAy(J), Avg(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    End If
Next
End Property

Private Property Get WEr_Tot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As WEr_
'Dim O As New WEr_
'Dim J%, C$, Ix%, Msg$
'For J = 0 To UB(Cnt)
'    C = Cnt(J)
'    Ix = Ay_Ix(Sum, C)
'    If Ix >= 0 Then
'        Msg = QQ_Fmt(M_Fld_IsCnt_FndInSum, CntLxAy(J), Cnt(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    Else
'        Ix = Ay_Ix(Avg, C)
'        If Ix >= 0 Then
'            Msg = QQ_Fmt(M_Fld_IsCnt_FndInAvg, CntLxAy(J), Cnt(J), AvgLxAy(Ix))
'            O.PushMsg Msg
'        End If
'    End If
'Next
'XSet_ WEr_Tot_1 = O
End Property

Private Property Get WEr_Wdt() As String()
End Property

Private Property Get WAny_Tot() As Boolean
Dim Lc As ListColumn
For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_XWAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
Next
End Property

Function Lo_Fmt(Lo As ListObject, Lo_Fmtr$()) As ListObject
Set A_Lo = Lo
A_Fny = Lo_Fny(Lo)
A_Lo_Fmtr = Lo_Fmtr

Bdr = X("Bdr")
Align = X("Align")
Tot = X("Tot")

Wdt = X("Wdt")
Fmt = X("Fmt")
Lvl = X("Lvl")
Cor = X("Cor")

Bet = X("Bet")
Fml = X("Fml")
Lbl = X("Lbl")
Tit = X("Tit")

Er_XHalt CvSy(AyAp_XAdd( _
    WEr_Align, WEr_Bdr, WEr_Tot, _
    WEr_Wdt, WEr_Fmt, WEr_Lvl, WEr_Cor, _
    WEr_Fml, WEr_Lbl, WEr_Tit, WEr_Bet))

XSet_Align
XSet_Bdr
XSet_Tot

XSet_Wdt
XSet_Lvl
XSet_Cor
XSet_Fmt

XSet_Tit
XSet_Lbl
XSet_Fml
XSet_Bet
End Function

Private Function WMsg_AlignLin(Ly$()) As String()
If Sz(Ly) Then Exit Function
End Function

Private Function WMsg_BetTermCnt$(L, NTerm%)

End Function

Private Function WMsg_Dup1(N, Ly$()) As String()
Dim L
For Each L In Ly
    If Lin_T1(L) = N Then PushI WMsg_Dup1, QQ_Fmt(M_Dup, L, N)
Next
End Function

Private Function WMsg_Dup(DupNy$(), Ly$()) As String()
Dim N
For Each N In AyNz(DupNy)
    PushIAy WMsg_Dup, WMsg_Dup1(N, Ly)
Next
End Function

Private Function WMsg_Fny(Fny$(), Lin_Ty$) As String()
'Return Msg if given-Fny has some field not in A_Fny
Dim WEr_Fny$(): WEr_Fny = AyMinus(Fny, A_Fny)
If Sz(WEr_Fny) = 0 Then Exit Function
PushI WMsg_Fny, QQ_Fmt(M_Fny, WEr_Fny, Lin_Ty)
End Function

Sub XBrw_Samp_LoFmtrTp()
Brw Samp_LoFmtrTp
End Sub

Private Function XSet_Align1(Fldss, A As XlHAlign)
Dim F
For Each F In A_Fny
    If XLik_Likss(F, Fldss) Then Lc_XSet_Align A_Lo, F, A
Next
End Function

Private Sub XSet_Align()
XSet_Align1 Ay_FstRmvT1(Align, "Left"), xlHAlignLeft
XSet_Align1 Ay_FstRmvT1(Align, "Right"), xlHAlignRight
XSet_Align1 Ay_FstRmvT1(Align, "Center"), xlHAlignCenter
End Sub

Private Sub XSet_Bdr()
Dim L$(), R$(), C$()
L = Ssl_Sy(JnSpc(Ay_XWh_RmvT1(Bdr, "Left")))
R = Ssl_Sy(JnSpc(Ay_XWh_RmvT1(Bdr, "Right")))
C = Ssl_Sy(JnSpc(Ay_XWh_RmvT1(Bdr, "Center")))
XSet_BdrLeft L
XSet_BdrLeft C
XSet_BdrRight C
XSet_BdrRight R
End Sub

Private Function XSet_BdrLeft(FldLikAy$())
Dim F
For Each F In A_Fny
    If XLik_LikAy(F, FldLikAy) Then Lc_XSet_BdrLeft A_Lo, F
Next
End Function

Private Function XSet_BdrRight(FldLikAy$())
Dim F
For Each F In A_Fny
    If XLik_LikAy(F, FldLikAy) Then Lc_XSet_BdrRight A_Lo, F
Next
End Function

Private Sub XSet_Bet()
Dim L, C$, X$, Y$
For Each L In AyNz(Tot)
    Lin_2TRstAsg L, C, X, Y
    A_Lo.ListColumns(C).DataBodyRange.Formula = QQ_Fmt("=Sum([?]:[?])", X, Y)
Next
End Sub

Private Sub XSet_Cor1(A)
Dim Cor1&, Fldss$, F
Lin_TRstAsg A, Cor1, Fldss
For Each F In A_Fny
    If XLik_Likss(F, Fldss) Then Lc_XSet_Cor A_Lo, F, Cor1
Next
End Sub

Private Sub XSet_Cor()
Dim L
For Each L In AyNz(Cor)
    XSet_Cor1 L
Next
End Sub

Private Sub XSet_Fml()
XSet_FmlBet

Dim C$, L, Fml1$
For Each L In AyNz(Fml)
    Lin_TRstAsg L, C, Fml1
    Lc_XSet_Fml A_Lo, C, Fml1
Next
End Sub

Private Sub XSet_FmlBet()

End Sub

Private Sub XSet_Lbl()

End Sub

Private Sub XSet_Lvl1(L)
Dim F
Dim Lvl As Byte, Fldss$
For Each F In A_Fny
    Lin_TRstAsg L, Lvl, Fldss
    If XLik_Likss(F, Fldss) Then
        Lc_XSet_Lvl A_Lo, F, Lvl
        Exit Sub
    End If
Next
End Sub

Private Sub XSet_Lvl()
Dim L
For Each L In AyNz(Lvl)
    XSet_Lvl1 L
Next
End Sub

Private Sub XSet_Tit()
Lo_XSet_Tit A_Lo, Tit
End Sub

Private Sub XSet_Tot1(FldLikss$, B As XlTotalsCalculation)
Dim F
For Each F In A_Fny
    If XLik_Likss(F, FldLikss) Then Lc_XSet_Tot A_Lo, F, B
Next
End Sub

Private Sub XSet_Tot()
XSet_Tot1 Ay_FstRmvT1(Tot, "Sum"), xlTotalsCalculationSum
XSet_Tot1 Ay_FstRmvT1(Tot, "Cnt"), xlTotalsCalculationCount
XSet_Tot1 Ay_FstRmvT1(Tot, "Avg"), xlTotalsCalculationAverage
End Sub

Private Sub XSet_Wdt1(L)
Dim F, W%, Likss$
Lin_TRstAsg L, W, Likss
For Each F In A_Fny
    If XLik_Likss(F, Likss) Then Lc_XSet_Wdt A_Lo, F, W: Exit For
Next
End Sub

Private Sub XSet_Wdt()
Dim L, F, W%, Likss1$
For Each L In AyNz(Wdt)
    XSet_Wdt1 L
Next
End Sub

Function Lo_LoFmtr(A As ListObject) As String()
Stop '
End Function
Sub Wb_Fmt_Lo(A As Workbook)
Dim I, L As ListObject
For Each I In Wb_LoAy(A)
    Set L = CvLo(I)
    Lo_Fmt L, Lo_LoFmtr(L)
Next
End Sub

Private Function X(T1$) As String()
X = Ay_XWh_RmvT1(A_Lo_Fmtr, T1)
End Function

Private Sub Z_WEr_Bet()
'---------------
A_Fny = Ssl_Sy("A B")
Erase Bet
    PushI Bet, "A B C"
    PushI Bet, "A B C"
Ept = EmpSy
    PushIAy Ept, WMsg_Dup(Ap_Sy("A"), Bet)
GoSub Tst
Exit Sub
'---------------
Tst:
    Act = WEr_Bet
    C
    Return
End Sub

Private Sub Z_Lo_Fmt()
Dim Lo As ListObject, LoFmtr$()
'------------
Set Lo = SampLo
LoFmtr = SampLoFmtr
GoSub Tst
Exit Sub
Tst:
    Lo_Fmt Lo, LoFmtr
    Return
End Sub

Private Sub Z_XSet_Bdr()
'--
Set A_Lo = SampLo
'--
Erase Bdr
PushI Bdr, "Left A B C"
PushI Bdr, "Left D E F"
PushI Bdr, "Right A B C"
PushI Bdr, "Center A B C"
GoSub Tst
Tst:
    
    XSet_Bdr      '<==
    Return
End Sub

Private Sub ZZ()
Dim A As ListObject
Dim B$()
Dim C As Workbook
Dim XX
Lo_Fmt A, B
XBrw_Samp_LoFmtrTp
Wb_Fmt_Lo C
End Sub

Private Sub Z()
Z_WEr_Bet
Z_Lo_Fmt
Z_XSet_Bdr
End Sub
