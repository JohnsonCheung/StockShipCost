Attribute VB_Name = "MTp_Sq_Sw"
Option Compare Binary
Option Explicit
Const CMod$ = "MTp_Sq_Sw."
Private A_Pm As Dictionary
Public Type SwRslt
    StmtSw As New Dictionary
    FldSw As Dictionary
    Er() As String
End Type

Private Sub AAMain()
Z_LnxAy_SwRslt
End Sub

Private Sub Evl(A() As SwBrk, OStmt As Dictionary, OFld As Dictionary)
Dim A1() As SwBrk
Dim A2() As SwBrk
Dim Sw As New Dictionary
Dim OSw As New Dictionary
Dim IsEvl As Boolean
Dim I%, J%
IsEvl = True
A1 = A
While IsEvl
    IsEvl = False
    J = J + 1
    If J > 1000 Then Stop
    For I = 0 To UB(A1)
        If SwBrk_EvlRslt(A1(I), Sw, OSw) Then
            IsEvl = True
        Else
            PushObj A2, A1(I)
        End If
    Next
    If False Then
        Brw AyAp_XAdd( _
            LblTabAy_Fmt("*A1", SwBrkAy_Fmt(A1)), _
            LblTabAy_Fmt("*A2", SwBrkAy_Fmt(A2)), _
            LblTabAy_Fmt("OSw", Dic_Fmt(OSw)))
        Stop
    End If
    A1 = A2
    Erase A2
    Set Sw = DicClone(OSw)
Wend
If Sz(OSw) > 0 Then Stop
EvlSwAsg OSw, OStmt, OFld
End Sub

Private Function EvlBoolTerm(A, Sw As Dictionary, ORslt As Boolean) As Boolean
If A_Pm.Exists(A) Then
    ORslt = A_Pm(A)
    EvlBoolTerm = True
    Exit Function
End If
If Not Sw.Exists(A) Then Exit Function
ORslt = Sw(A)
EvlBoolTerm = True
End Function

Private Function SwBrk_EvlRslt(A As SwBrk, Sw As Dictionary, OSw As Dictionary) As Boolean
'Return True and set Result if evalulated
Const CSub$ = CMod & "SwBrk_EvlRslt"
If Sw.Exists(A.Nm) Then XThw CSub, "[Lnx_SwBrk] should not be found in [Sw]", "", SwBrk_Lin(A), Dic_Fmt(Sw)
Dim Ay$(): Ay = A.TermAy
Dim ORslt As Boolean, IsEvl As Boolean
Select Case A.OpStr
Case "OR":  IsEvl = EvlTermAy(Ay, "OR", Sw, ORslt)
Case "AND": IsEvl = EvlTermAy(Ay, "AND", Sw, ORslt)
Case "NE":  IsEvl = EvlT1T2(Ay(0), Ay(1), "NE", Sw, ORslt)
Case "EQ":  IsEvl = EvlT1T2(Ay(0), Ay(1), "EQ", Sw, ORslt)
Case Else: XThw CSub, "Invalid Op", "Invalid-Op In-SwBrk-Lin Valid-OpStr", A.OpStr, SwBrk_Lin(A), "OR AND NE EQ"
End Select
If IsEvl Then
    OSw.Add A.Nm, ORslt
    SwBrk_EvlRslt = True
End If
End Function

Private Sub EvlSwAsg(A As Dictionary, OStmt As Dictionary, OFld As Dictionary)
Set OStmt = New Dictionary
Set OFld = New Dictionary
Dim K
For Each K In A.Keys
    If XTak_FstTwoChr(K) <> "?#" Then  ' Skip.  It is temp-Sw
        Select Case Left(K, 5)
        Case "?SEL#", "?UPD#"     ' It is StmtSw
            OStmt.Add Mid(K, 5), A(K)
        Case Else                 ' It is FldSw
            OFld.Add K, A(K)
        End Select
    End If
Next
End Sub

Private Function EvlT1(A, Sw As Dictionary, ORslt$) As Boolean
EvlT1 = EvlTerm(A, Sw, ORslt)
End Function

Private Function EvlT1T2(T1$, T2$, EQ_NE$, Sw As Dictionary, ORslt As Boolean) As Boolean
'Return True and set ORslt if evaluated
Const CSub$ = CMod & "EvlT1T2"
Dim S1$, S2$
If Not EvlT1(T1, Sw, S1) Then Exit Function
If Not EvlT2(T2, Sw, S2) Then Exit Function
Select Case EQ_NE
Case "EQ": ORslt = S1 = S2: EvlT1T2 = True
Case "NE": ORslt = S1 <> S2:: EvlT1T2 = True
Case Else: XThw CSub, "Parameter [EQ_NE] should be EQ or NE", "EQ_NE", EQ_NE
End Select
End Function

Private Function EvlT2(A, Sw As Dictionary, ORslt$) As Boolean
'Return True is evalulated
'switch-term begins with @ or ? or it is *Blank.  @ is for parameter & ? is for switch
'  If @, it will evaluated to str by lookup from Pm
'        if not exist in {Pm}, stop, it means the validation fail to remove this term
'  If ?, it will evaluated to bool by lookup from Sw
'        if not exist in {Sw}, return None
'  Otherwise, just return SomVar(A)
If EvlTerm(A, Sw, ORslt) Then Exit Function
If XTak_FstChr(A) = "*" Then
    If UCase(A) <> "*BLANK" Then Stop ' it means the validation fail to remove this term
    ORslt = ""
Else
    ORslt = A
End If
EvlT2 = True
End Function

Private Function EvlTerm(A, Sw As Dictionary, ORslt$) As Boolean
If A_Pm.Exists(A) Then
    ORslt = A_Pm(A)
    EvlTerm = True
    Exit Function
End If
If Not Sw.Exists(A) Then Exit Function
ORslt = Sw(A)
EvlTerm = True
End Function

Private Function EvlTermAy1(A$(), AND_OR, Sw As Dictionary) As Boolean()
Dim Rslt As Boolean, O() As Boolean, I
For Each I In A
    If Not EvlBoolTerm(I, Sw, Rslt) Then Exit Function
    PushI O, Rslt
Next
EvlTermAy1 = O
End Function

Private Function EvlTermAy(A$(), AND_OR$, Sw As Dictionary, ORslt As Boolean) As Boolean
If Sz(A) = 0 Then Stop
Dim BoolAy() As Boolean
    BoolAy = EvlTermAy1(A, AND_OR, Sw)
    If Sz(BoolAy) = 0 Then Exit Function
    
Select Case AND_OR
Case "AND": ORslt = BoolAy_IsAllTrue(BoolAy)
Case "OR":  ORslt = BoolAy_IsSomTrue(BoolAy)
Case Else: Stop
End Select
EvlTermAy = True
End Function

Function LnxAy_SwRslt(A() As Lnx, Pm As Dictionary) As SwRslt
Set A_Pm = Pm
Dim B() As SwBrk
Dim O As SwRslt
AyAsg WEr(A), O.Er, B
Evl B, O.StmtSw, O.FldSw
LnxAy_SwRslt = O
End Function

Private Property Get Samp_Pm() As Dictionary
Set Samp_Pm = New_Dic_LY(Samp_PmLy)
End Property

Private Property Get Samp_PmLy() As String()
PushI Samp_PmLy, "@?BrkMbr 0"
PushI Samp_PmLy, "@?BrkSto 0"
PushI Samp_PmLy, "@?BrkCrd 0"
PushI Samp_PmLy, "@?BrkDiv 0"
'-- @XXX means txt and optional, allow, blank
PushI Samp_PmLy, "@SumLvl  Y"
PushI Samp_PmLy, "@?MbrEmail 1"
PushI Samp_PmLy, "@?MbrNm    1"
PushI Samp_PmLy, "@?MbrPhone 1"
PushI Samp_PmLy, "@?MbrAdr   1"
'-- @@ mean compulasary
PushI Samp_PmLy, "@DteFm 20170101"
'@@DteTo 20170131
PushI Samp_PmLy, "@LisDiv 1 2"
PushI Samp_PmLy, "@LisSto"
PushI Samp_PmLy, "@LisCrd"
PushI Samp_PmLy, "@CrdExpr ..."
PushI Samp_PmLy, "@CrdExpr ..."
PushI Samp_PmLy, "@CrdExpr ..."
End Property

Private Property Get Samp_SwLnxAy() As Lnx()
Dim M$()
PushI M, "?#LvlY    EQ @SumLvl Y"
PushI M, "?#LvlY    EQ @SumLvl Y"
PushI M, "?#LvlM    EQ @SumLvl M"
PushI M, "?#LvlW    EQ @SumLvl W"
PushI M, "?#LvlD    EQ @SumLvl D"
PushI M, "?Y       OR ?#LvlD ?#LvlW ?#LvlM ?#LvlY"
PushI M, "?M       OR ?#LvlD ?#LvlW ?#LvlM"
PushI M, "?W       OR ?#LvlD ?#LvlW"
PushI M, "?D       OR ?#LvlD"
PushI M, "?Dte     OR ?#LvlD"
PushI M, "?Mbr     OR @?BrkMbr XX"
PushI M, "?MbrCnt  OR @?BrkMbr"
PushI M, "?Div     OR @?BrkDiv"
PushI M, "?Sto     OR @?BrkSto"
PushI M, "?Crd     OR @?BrkCrd"
PushI M, "?SEL#Div NE @LisDiv *blank"
PushI M, "?SEL#Sto NE @LisSto *blank"
PushI M, "?SEL#Crd NE @LisCrd *blank"
Samp_SwLnxAy = Ly_LnxAy(M)
End Property

Private Function WEr(A() As Lnx) As Variant()
Dim B() As SwBrk
    B = LnxAy_SwBrkAy(A)
'    Er = AyAp_XAdd(XChk_Lin(B), XChk_DupNm(B), XChk_Fld(B), XChk_LeftOvr(B))
'WEr = Array(Er, B)
End Function

Private Function SwBrkAy_Fmt(A() As SwBrk) As String()
Dim I
For Each I In AyNz(A)
    PushI SwBrkAy_Fmt, SwBrk_Lin(CvSwBrk(I))
Next
End Function

Private Function SwBrk_Msg$(A As SwBrk, B$)
SwBrk_Msg = SwBrk_Lin(A) & " --- " & B
End Function

Private Function SwBrk_MsgDupNm(A() As SwBrk) As String()
Dim I
For Each I In AyNz(A)
    PushI SwBrk_MsgDupNm, SwBrk_Msg(CvSwBrk(I), "Dup name")
Next
End Function

Private Function SwBrk_MsgLeftOvrAftEvl(A() As SwBrk, Sw As Dictionary) As String()
If Sz(A) = 0 Then Exit Function
Dim I
PushI SwBrk_MsgLeftOvrAftEvl, "Following lines cannot be further evaluated:"
For Each I In A
    PushI SwBrk_MsgLeftOvrAftEvl, vbTab & SwBrk_Lin(CvSwBrk(I))
Next
PushIAy SwBrk_MsgLeftOvrAftEvl, DicLblLy(Sw, "Following is the [Sw] after evaluated:")
End Function

Private Function SwBrk_MsgNoNm$(A As SwBrk)
SwBrk_MsgNoNm = SwBrk_Msg(A, "No name")
End Function

Private Function SwBrk_MsgOpStrEr$(A As SwBrk)
SwBrk_MsgOpStrEr = SwBrk_Msg(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function

Private Function SwBrk_MsgPfx$(A As SwBrk)
SwBrk_MsgPfx = SwBrk_Msg(A, "First Char must be @")
End Function

Private Function SwBrk_MsgTermCntAndOr$(A As SwBrk)
SwBrk_MsgTermCntAndOr = SwBrk_Msg(A, "When 2nd-Term (Operator) is [AND OR], at least 1 term")
End Function

Private Function SwBrk_MsgTermCntEqNe$(A As SwBrk)
SwBrk_MsgTermCntEqNe = SwBrk_Msg(A, "When 2nd-Term (Operator) is [EQ NE], only 2 terms are allowed")
End Function

Private Function SwBrk_MsgTermMustBegWithQuestOrAt$(TermAy$(), A As SwBrk)
SwBrk_MsgTermMustBegWithQuestOrAt = SwBrk_Msg(A, "Terms[" & JnSpc(TermAy) & "] must begin with either [?] or [@?]")
End Function

Private Function SwBrk_MsgTermNotInPm$(TermAy$(), A As SwBrk)
SwBrk_MsgTermNotInPm = SwBrk_Msg(A, "Terms[" & JnSpc(TermAy) & "] begin with [@?] must be found in Pm")
End Function

Private Function SwBrk_MsgTermNotInSw$(TermAy$(), A As SwBrk, SwNm As Dictionary)
SwBrk_MsgTermNotInSw = SwBrk_Msg(A, "Terms[" & JnSpc(TermAy) & "] begin with [?] must be found in Switch")
End Function

Private Property Get WOpStrAy() As String()
Static X$()
If Sz(X) = 0 Then X = Ssl_Sy("OR AND NE EQ")
WOpStrAy = X
End Property

Private Property Get WPmNy() As String()
WPmNy = DicStrKy(A_Pm)
End Property

Private Function Lnx_SwBrk(A As Lnx) As SwBrk
Const CSub$ = CMod & "Lnx_SwBrk"
Dim L$, Ix%, OEr$()
L = A.Lin
Ix = A.Ix
If Lin_IsRmkLin(L) Then XThw CSub, "[SwLin], [Ix] is a remark line.  It should be removed before calling Evl", A.Lin, A.Ix
Set Lnx_SwBrk = New SwBrk
With Lnx_SwBrk
    .Nm = XShf_Term(L)
    .OpStr = UCase(XShf_Term(L))
    .TermAy = Ssl_Sy(L)
    .Ix = Ix
End With
End Function

Private Function LnxAy_SwBrkAy(A() As Lnx) As SwBrk()
Dim I
For Each I In AyNz(A)
    PushObj LnxAy_SwBrkAy, Lnx_SwBrk(CvLnx(I))
Next
End Function

Private Function SwBrk_Lin$(A As SwBrk)
With A
    SwBrk_Lin = XQuote(.Ix, "L#(*) ") & XQuote_SqBkt(JnSpc(Array(.Nm, .OpStr, JnSpc(.TermAy))))
End With
End Function

Private Function WSwNmDic(A() As SwBrk) As Dictionary
Set WSwNmDic = AyIxDic(WSwNy(A))
End Function

Private Function WSwNy(A() As SwBrk) As String()
Dim J%
For J = 0 To UB(A)
    PushI WSwNy, A(J).Nm
Next
End Function

Private Sub SwBrkAy_XBrw(A() As SwBrk)
Ay_XBrw SwBrkAy_Fmt(A)
End Sub

Private Function XChk_DupNm(IO() As SwBrk) As String()
Dim Ny$(), Nm$, O() As SwBrk
Dim J%, M As SwBrk, Er() As SwBrk
For J = 0 To UB(IO)
    Set M = IO(J)
    If Ay_XHas(Ny, M.Nm) Then
        PushObj Er, M
    Else
        PushObj O, M
        PushI Ny, M.Nm
    End If
Next
XChk_DupNm = SwBrk_MsgDupNm(Er)
IO = O
End Function

Private Function XChk_Fld(A() As SwBrk, OEr$()) As SwBrk()
XChk_Fld = A
Exit Function
Dim M As SwBrk, IsEr As Boolean, J%, I, A1() As SwBrk, A2() As SwBrk
IsEr = True
A1 = A
While IsEr
    J = J + 1: If J > 1000 Then Stop
    IsEr = False
    For Each I In AyNz(A1)
        Set M = XChk_FldLin(CvSwBrk(I), WSwNmDic(A), OEr)
        If IsNothing(M) Then
            IsEr = True
        Else
            PushObj A2, M
        End If
    Next
    A1 = A2
Wend
XChk_Fld = A2
End Function

Private Function XChk_FldLin(A As SwBrk, SwNm As Dictionary, OEr$()) As SwBrk
'Each Term in A.TermAy must be found either in Sw or Pm
Dim O0$(), O1$(), O2$(), I
For Each I In AyNz(A.TermAy)
    Select Case True
    Case XHas_Pfx(I, "?"):  If Not SwNm.Exists(I) Then Push O0, I
    Case XHas_Pfx(I, "@?"): If Not A_Pm.Exists(I) Then Push O1, I
    Case Else:                  Push O2, I
    End Select
Next
PushIAy OEr, SwBrk_MsgTermNotInSw(O0, A, SwNm)
PushIAy OEr, SwBrk_MsgTermNotInPm(O1, A)
PushIAy OEr, SwBrk_MsgTermMustBegWithQuestOrAt(O2, A)
If AyApHasEle(O0, O1, O2) Then Set XChk_FldLin = A
End Function

Function XChk_LeftOvr(IO() As SwBrk) As String()
End Function

Private Function XChk_Lin1$(IO As SwBrk)
Dim SwBrk_Msg$
With IO
    If .Nm = "" Then XChk_Lin1 = SwBrk_MsgNoNm(IO): Exit Function
    Select Case .OpStr
    Case "OR", "AND": If Sz(.TermAy) = 0 Then XChk_Lin1 = SwBrk_MsgTermCntAndOr(IO): Exit Function
    Case "EQ", "NE":  If Sz(.TermAy) <> 2 Then XChk_Lin1 = SwBrk_MsgTermCntEqNe(IO): Exit Function
    Case Else:        XChk_Lin1 = SwBrk_MsgOpStrEr(IO): Exit Function
    End Select
End With
End Function

Private Function XChk_Lin(IO() As SwBrk) As String()
Dim I
For Each I In AyNz(IO)
    PushNonNothing XChk_Lin, XChk_Lin1(CvSwBrk(I))
Next
End Function

Private Function XChk_Pfx(A() As Lnx, OEr$()) As Lnx()
Dim J%
For J = 0 To UB(A)
    If XTak_FstChr(A(J).Lin) <> "?" Then
        PushI OEr, SwBrk_MsgPfx(A(J))
    Else
        PushObj XChk_Pfx, A(J)
    End If
Next
End Function

Private Sub Z_LnxAy_SwRslt()
Dim Stmt As Dictionary, Fld As Dictionary, Er$()
Dim SR As SwRslt
SR = LnxAy_SwRslt(Samp_SwLnxAy, Samp_Pm)
Brw NyAp_Ly("SwLnxAy Pm StmtSw FldSw", _
    LnxAy_Fmt(Samp_SwLnxAy), _
    A_Pm, _
    SR.StmtSw, _
    SR.FldSw)
End Sub

Private Sub ZZ()
Dim A() As Lnx
Dim B As Dictionary
Dim C() As SwBrk
Dim XX
LnxAy_SwRslt A, B
XChk_LeftOvr C
End Sub

Private Sub Z()
Z_LnxAy_SwRslt
End Sub
