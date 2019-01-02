Attribute VB_Name = "MDao_Schm_Er"
Option Compare Binary
Option Explicit
Const MD_NTermShouldBe3OrMore$ = "Should have 3 or more terms"

Private Sub AMain()
Dim X$()
SchmLy_Er X
End Sub

Private Function ErOf_D_DLnxAy(A() As Lnx) As String()
Dim I
For Each I In AyNz(A)
    If Lin_2TRst(CvLnx(I).Lin)(2) = "" Then PushI ErOf_D_DLnxAy, MsgOf_D_NTermShouldBe3OrMore(CvLnx(I))
Next
End Function

Private Function ErOf_D_F_FnyNotIn_Fny(D() As Lnx, T$, Fny$()) As String()
'Fny$ is the fields in T$.
'D-Lnx.Lin is Des $T $F $D.  For the T$=$T line, the $F not in Fny$(), it is error
Dim I
For Each I In AyNz(D)
    PushNonBlankStr ErOf_D_F_FnyNotIn_Fny, ErOf_D_F_FnyNotIn_Fny1(CvLnx(I), T, Fny)
Next
End Function

Private Function ErOf_D_F_FnyNotIn_Fny1$(D As Lnx, T$, Fny$())
Dim Tbl$, Fld$, Des$, X$
Lin_3TRstAsg D.Lin, X, Tbl, Fld, Des
If X <> "Des" Then Stop
If Tbl <> T Then Exit Function
If Not Ay_XHas(Fny, Fld) Then
    ErOf_D_F_FnyNotIn_Fny1 = MsgOf_D_FldNotIn_Fny()
End If
End Function

Private Function ErOf_D_FldEr(A() As Lnx, T() As Lnx) As String() _
'Given A is D-Lnx having fmt = $Tbl $Fld $D, _
'This Sub checks if $Fld is in Fny
Dim J%, Fld$, Tbl$, Tny$(), FnyAy$()

For J = 0 To UB(A)
    Lin_TTAsg A(J).Lin, Tbl, Fld
    If Tbl <> "." Then
        PushIAy ErOf_D_FldEr, ErOf_D_FldEr1(A(J), Tny, FnyAy)
    End If
Next
End Function

Private Function ErOf_D_FldEr1(A As Lnx, Tny$(), FnyAy$()) As String()

End Function

Private Function ErOf_D_T_Tbl_NotIn_Tny1$(D As Lnx, Tny$())

End Function

Private Function ErOf_D_TblEr(D() As Lnx, Tny$()) As String()
'D-Lnx.Lin is Des $T $F $D. If $T<>"." and not in Tny, it is error
Dim J%
For J = 0 To UB(D)
    PushNonBlankStr ErOf_D_TblEr, ErOf_D_T_Tbl_NotIn_Tny1(D(J), Tny)
Next
End Function

Private Function ErOf_E_DupE(E() As Lnx, Eny$()) As String()
Dim Ele
For Each Ele In AyNz(Ay_XWh_Dup(Eny))
    Push ErOf_E_DupE, MsgOf_E_DupE(LnxAyEle_LnoAy(E, Ele), Ele)
Next
End Function

Private Function ErOf_E_ELnx(A As Lnx) As String()
Dim L$, Ty$
Dim E$  ' EleNm of the Ele-Lin
L = A.Lin
'    AyAsg XShf_Val(L, "*Nm *Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
'                     .E, Ty, .Req, .ZLen, .TxtSz, .VTxt, .Dft, .VRul, .Expr
'    .Ty = DaoTy(Ty)
    If L <> "" Then
        Dim ExcessEle$
        Push ErOf_E_ELnx, MsgOf_E_ExcessEleItm(A, ExcessEle)
    End If
'    If .Ty = 0 Then
'        Push OEr, MsgOf_TyEr(A.Ix, Ty)
'    End If
End Function

Private Function ErOf_E_ELnxAy(A() As Lnx) As String()

End Function

Private Function ErOf_T_FldHasNoEle(T() As Lnx, E() As Lnx) As String()

End Function

Private Function ErOf_F_Ele_NotIn_Eny(F() As Lnx, Eny$()) As String()
Dim J%, O$(), Eless$, E$
For J = 0 To UB(F)
    With F(J)
        E = Lin_T1(F(J).Lin)
        If Not Ay_XHas(Eny, E) Then PushI ErOf_F_Ele_NotIn_Eny, MsgOf_F_Ele_NotIn_Eny(F(J), E, Eless)
    End With
Next
ErOf_F_Ele_NotIn_Eny = O
End Function

Private Function ErOf_F_EleHasNoDef(F() As Lnx, AllEny$()) As String() _

End Function

Private Function ErOf_F_FLnx(F As Lnx) As String()
Dim LikFF$, A$, V$
'    AyAsg Lin_3TRst(F.Lin), .E, .LikT, V, A
'    .LikFny = Ssl_Sy(LikFF)
End Function

Private Function ErOf_F_FLnxAy(A() As Lnx) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy ErOf_F_FLnxAy, ErOf_F_FLnx(A(J))
Next
End Function

Private Function ErOf_T_DupT(T() As Lnx, Tny$()) As String()
Dim Tbl
For Each Tbl In AyNz(Ay_XWh_Dup(Tny))
    Push ErOf_T_DupT, MsgOf_T_DupT(LnxAyTbl_LnoAy(T, Tbl), Tbl)
Next
End Function

Private Function ErOf_T_NoTLin(A() As Lnx) As String()
If Sz(A) > 0 Then Exit Function
PushI ErOf_T_NoTLin, MsgOf_T_NoTLin
End Function

Private Function ErOf_T_TLnx(T As Lnx) As String()
Dim L$
Dim Tbl$
    L = T.Lin
    Tbl = XShf_Term(L)
    L = Replace(L, "*", Tbl)
'1
Select Case SubStrCnt(L, "|")
Case 0, 2
Case Else: PushI ErOf_T_TLnx, MsgOf_T_VBar_Cnt(T): Exit Function
End Select

'2
If Not IsNm(Tbl) Then
    PushI ErOf_T_TLnx, MsgOf_T_TblIsNotNm(T)
    Exit Function
End If
'
Dim Fny$()
    Fny = Ssl_Sy(Replace(L, "|", " "))
    
If HasSubStr(L, "|") Then
'3
    Dim IdFld$
    IdFld = Trim(XTak_Bef(L, "|"))
    If IdFld <> Tbl & "Id" Then
        PushI ErOf_T_TLnx, MsgOf_T_IdFld(T)
        Exit Function
    End If
'4
    If Trim(TakBet(L, "|", "|")) = "" Then
        PushI ErOf_T_TLnx, MsgOf_T_NoFLdBetVV(T)
        Exit Function
    End If
End If
'5
    Dim Dup$()
    Dup = Ay_XWh_Dup(Fny)
    If Sz(Dup) > 0 Then
        PushI ErOf_T_TLnx, MsgOf_T_DupF(T, Tbl, Dup)
        Exit Function
    End If
'6
If Sz(Fny) = 0 Then
    PushI ErOf_T_TLnx, MsgOf_T_NoFld(T)
    Exit Function
End If
'7
Dim F
For Each F In AyNz(Fny)
    If Not IsNm(F) Then
        PushI ErOf_T_TLnx, MsgOf_T_FldIsNotANmEr(T, F)
    End If
Next
End Function

Private Function ErOf_T_TLnxAy(A() As Lnx) As String()
Dim I
For Each I In AyNz(A)
    PushIAy ErOf_T_TLnxAy, ErOf_T_TLnx(CvLnx(I))
Next
End Function

Private Function LnxAyEle_LnoAy(E() As Lnx, Ele) As Long()
Dim J%
For J = 0 To UBound(E)
    If Lin_T1(E(J).Lin) = Ele Then
        PushI LnxAyEle_LnoAy, E(J).Ix + 1
    End If
Next
End Function

Private Function LnxAyTbl_LnoAy(A() As Lnx, T) As Long()
Dim J%
For J = 0 To UB(A)
    If Lin_T1(A(J).Lin) = T Then
        PushI LnxAyTbl_LnoAy, A(J).Ix + 1
    End If
Next
End Function

Private Function LnoAyMsgOf__MsgOf_$(A&(), M$)
LnoAyMsgOf__MsgOf_ = QQ_Fmt("--- #[?] ?", JnSpc(A), M)
End Function

Private Function MsgOf_LnxMsg$(A As Lnx, M$)
MsgOf_LnxMsg = QQ_Fmt("--- #?[?] ?", A.Ix + 1, A.Lin, M)
End Function

Private Property Get MsgOf_D_FldNotIn_Fny$()

End Property

Private Function MsgOf_D_NTermShouldBe3OrMore$(D As Lnx)
MsgOf_D_NTermShouldBe3OrMore = MsgOf_LnxMsg(D, MD_NTermShouldBe3OrMore)
End Function

Private Function MsgOf_D_T_Tbl_NotIn_Tny$(A As Lnx, T, Tblss$)
MsgOf_D_T_Tbl_NotIn_Tny = MsgOf_LnxMsg(A, QQ_Fmt("T[?] is invalid.  Valid T[?]", T, Tblss))
End Function

Private Function MsgOf_E_DupE$(LnoAy&(), E)
MsgOf_E_DupE = LnoAyMsgOf__MsgOf_(LnoAy, QQ_Fmt("This E[?] is dup", E))
End Function

Private Function MsgOf_T_DupT$(LnoAy&(), T)
MsgOf_T_DupT = LnoAyMsgOf__MsgOf_(LnoAy, QQ_Fmt("This Tbl[?] is dup", T))
End Function

Private Function MsgOf_E_ExcessEleItm$(A As Lnx, ExcessEle$)
MsgOf_E_ExcessEleItm = MsgOf_LnxMsg(A, QQ_Fmt("Excess Ele Item [?]", ExcessEle))
End Function

Private Function MsgOf_F_ExcessTxtSz$(A As Lnx)
MsgOf_F_ExcessTxtSz = MsgOf_LnxMsg(A, "Non-Txt-Ty should not have TxtSz")
End Function

Private Function MsgOf_F_Ele_NotIn_Eny$(A As Lnx, E$, Eless$)
MsgOf_F_Ele_NotIn_Eny = MsgOf_LnxMsg(A, QQ_Fmt("Ele of is not in F-Lin not in Eny", E, Eless))
End Function

Private Function MsgOf_E_FldEleEr$(A As Lnx, E$, Eless$)
MsgOf_E_FldEleEr = MsgOf_LnxMsg(A, QQ_Fmt("E[?] is invalid.  Valid E is [?]", E, Eless))
End Function

Private Function MsgOf_F_DLy_NotIn_Tbl_Fny$(A As Lnx, T$, F$, Fssl$)
MsgOf_F_DLy_NotIn_Tbl_Fny = MsgOf_LnxMsg(A, QQ_Fmt("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function MsgOf_T_DupF$(A As Lnx, T$, Fny$())
MsgOf_T_DupF = MsgOf_LnxMsg(A, QQ_Fmt("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function MsgOf_T_FldIsNotANmEr$(A As Lnx, F)
MsgOf_T_FldIsNotANmEr = MsgOf_LnxMsg(A, QQ_Fmt("FldNm[?] is not a name", F))
End Function

Private Function MsgOf_T_IdFld$(A As Lnx)
Const M$ = "The field before first | must be *Id field"
MsgOf_T_IdFld = MsgOf_LnxMsg(A, M)
End Function

Private Function MsgOf_T_NoFld(A As Lnx)
MsgOf_T_NoFld = MsgOf_LnxMsg(A, "No field")
End Function

Private Function MsgOf_T_NoFLdBetVV$(A As Lnx)
MsgOf_T_NoFLdBetVV = MsgOf_LnxMsg(A, "No field between | |")
End Function

Private Property Get MsgOf_T_NoTLin$()
MsgOf_T_NoTLin = "No T-Line"
End Property

Private Function MsgOf_T_TblIsNotNm$(A As Lnx)
MsgOf_T_TblIsNotNm = MsgOf_LnxMsg(A, "Tbl is not a name")
End Function

Private Function MsgOf_T_VBar_Cnt$(A As Lnx)
Const M$ = "The T-Lin should have 0 or 2 VBar only"
MsgOf_T_VBar_Cnt = MsgOf_LnxMsg(A, M)
End Function

Private Function MsgOf_T_FldEr$(A As Lnx, F$)
MsgOf_T_FldEr = MsgOf_LnxMsg(A, QQ_Fmt("Fld[?] cannot be found in any Ele-Lines"))
End Function

Private Function MsgOf_TyEr$(A As Lnx, Ty$)
MsgOf_TyEr = MsgOf_LnxMsg(A, QQ_Fmt("Invalid DaoTy[?].  Valid Ty[?]", Ty, DaoTySsl))
End Function

Private Property Get Samp_Schm$()
Const A_1$ = "Tbl $Lo_FmtWs  | WsNm |" & _
vbCrLf & "Tbl $Lo_FmtWdt          WsNm Seq Wdt ColNmss Er" & _
vbCrLf & "Tbl $Lo_FmtLvl | WsNm" & _
vbCrLf & "Tbl $Lnk" & _
vbCrLf & "Fld " & _
vbCrLf & "Fld $"

Samp_Schm = A_1
End Property

Function SchmLy_Er(A$()) As String()
'========================================
Dim E() As Lnx
Dim F() As Lnx
Dim D() As Lnx
Dim T() As Lnx
    Dim X()  As Lnx
    X = LyClnLnxAy(A)
    T = LnxAy_XWh_RmvT1(X, "Tbl")
    D = LnxAy_XWh_RmvT1(X, "Des")
    E = LnxAy_XWh_RmvT1(X, "Ele")
    F = LnxAy_XWh_RmvT1(X, "Fld")

Dim Er$(), Tny$(), Eny$()
    Eny = Ay_XTak_T1(LnxAy_Ly(E))
    Tny = Ay_XTak_T1(LnxAy_Ly(T))
    Er = LnxAyT1Chk(X, "Des Ele Fld Tbl")
'========================================
Dim AllFny$()
Dim AllEny$()
    AllFny = TdStrAy_Fny(LnxAy_Ly(T))
    AllEny = Ay_XTak_T1(LnxAy_Ly(E))
SchmLy_Er = AyAp_XAdd( _
Er, _
ErOf_T_TLnxAy(T), _
ErOf_F_FLnxAy(F), _
ErOf_E_ELnxAy(E), _
ErOf_D_DLnxAy(D), _
ErOf_T_DupT(T, Tny), _
ErOf_T_NoTLin(T), _
ErOf_E_DupE(E, Eny), _
ErOf_T_FldHasNoEle(T, E), _
ErOf_D_TblEr(D, Tny), _
ErOf_D_FldEr(D, T), _
ErOf_F_EleHasNoDef(F, AllEny))
End Function

Private Sub Z_ErOf_T_TLnx()
GoSub Cas0
Stop
GoSub Cas1
GoSub Cas2
GoSub Cas3
GoSub Cas4
GoSub Cas5
GoSub Cas6
Exit Sub
Dim EptEr$(), ActEr$()
Dim TLnx As New Lnx
Cas0:
    Set TLnx = Lnx(999, "Tbl 1")
    Ept = Ap_Sy("--- #1000[Tbl 1] FldNm[1] is not a name")
    GoTo Tst
Cas1:
    Set TLnx = Lnx(999, "A")
    Push EptEr, "should have a |"
    Ept = Ap_Sy("")
    GoTo Tst
Cas2:
    TLnx.Lin = "A | B B"
    Ept = Ap_Sy("")
    Push EptEr, "dup fields[B]"
    GoTo Tst
Cas3:
    TLnx.Lin = "A | B B D C C"
    Ept = Ap_Sy("")
    Push EptEr, "dup fields[B C]"
    GoTo Tst
Cas4:
    TLnx.Lin = "A | * B D C"
    Ept = Ap_Sy("")
    With Ept
        .T = "A"
        .Fny = Ssl_Sy("A B D C")
    End With
    GoTo Tst
Cas5:
    TLnx.Lin = "A | * B | D C"
    Ept = Ap_Sy("")
    With Ept
        .T = "A"
        .Fny = Ssl_Sy("A B D C")
        .Sk = Ssl_Sy("B")
    End With
    GoTo Tst
Cas6:
    TLnx.Lin = "A |"
    Ept = Ap_Sy("")
    Push EptEr, "should have fields after |"
    GoTo Tst
Tst:
    Act = ErOf_T_TLnx(TLnx)
    C
    Return
End Sub

Private Sub Z()
Z_ErOf_T_TLnx
Exit Sub
'AAAA
'SchmLyEr
End Sub
