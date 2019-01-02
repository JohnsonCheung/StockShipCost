Attribute VB_Name = "MSql_QSqp"
Option Compare Binary
Option Explicit
Const C_Pfx$ = vbTab
Const C_Upd$ = "update"
Const C_Into$ = "into"
Const C_Sel$ = "select"
Const C_SelDis$ = "select distinct"
Const C_Fm$ = "from"
Const C_Gp$ = "group by"
Const C_Wh$ = "where"
Const C_And$ = "and"
Const C_Jn$ = "join"
Const C_LeftJn$ = "left join"
Const C_NLT$ = vbCrLf & vbTab
Const C_NLTT$ = vbCrLf & vbTab & vbTab
Public IsSqlFmt As Boolean
Private Property Get C_NL$() ' New Line
If IsSqlFmt Then
    C_NL = vbCrLf
Else
    C_NL = " "
End If
End Property
Private Property Get NLT$() ' New Line Tabe
If IsSqlFmt Then
    NLT = C_NLT
Else
    NLT = " "
End If
End Property
Private Property Get NLTT$() ' New Line Tabe
If IsSqlFmt Then
    NLTT = C_NLTT
Else
    NLTT = " "
End If
End Property

Private Function AyQ(Ay) As String()
AyQ = AyQuote(Ay, Var_XQuote_Sql(Ay(0)))
End Function
Function FldInAy$(F, InAy)
FldInAy = Q(F) & "(" & JnComma(AyQ(InAy)) & ")"
End Function
Private Function EyFny_AsLines$(Ey$(), Fny$())
Stop '
End Function
Function FF_JnComma$(FF)
FF_JnComma = JnComma(Ay_XQuote_SqBkt_IfNeed(FF_Fny(FF)))
End Function
Function QQInto_T$(T)
QQInto_T = NLT & "Into" & NLTT & "[" & T & "]"
End Function

Function AddColSqp$(Fny0, FldDfnDic As Dictionary)
Dim Fny$(), O$(), J%
Fny = CvNy(Fny0)
ReDim O(UB(Fny))
For J = 0 To UB(Fny)
    O(J) = Fny(J) & " " & FldDfnDic(Fny(J))
Next
'AddColSqp = NxtLin & "Add Column " & JnComma(O)
End Function

Function BExpr_AndSqp$(A$)
If A = "" Then Exit Function
'AndSqp = NxtLin & "and " & NxtLin_Tab & Expr
End Function

Private Function Ay_XAdd_PfxNLTT$(A)
Ay_XAdd_PfxNLTT = vbCrLf & JnCrLf(Ay_XAdd_Pfx(A, C_Pfx & C_Pfx))
End Function

Function ExprInLis_InLisBExpr$(Expr$, InLis$)
If InLis = "" Then Exit Function
ExprInLis_InLisBExpr = QQ_Fmt("? in (?)", Expr, InLis)
End Function

Function QQFm_T$(T)
QQFm_T = NLT & C_Fm & NLTT & XQuote_SqBkt(T)
End Function

Function Gp$(ExprVblAy$())
Ass IsVblAy(ExprVblAy)
Gp = VblAy_AlignAsLines(ExprVblAy, "|  Group By")
End Function

Private Sub Z_Gp()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
Ay_XDmp SplitVBar(Gp(ExprVblAy))
End Sub

Function SqpSelFldLvs$(FldLvs$, ExprVblAy$())
Dim Fny$(): Fny = Ssl_Sy(FldLvs)
'SqpSelFldLvs = SqpSel(Fny, ExprVblAy)
End Function

Private Sub SqpSel()
Dim Fny$(), ExprVblAy$()
ExprVblAy = Ap_Sy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print XRpl_VBar(SqpSel(Fny, ExprVblAy))
End Sub

Function Into$(T)
Into = NLT & C_Into & NLTT & Q(T)
End Function
Function QQSel_X$(X)
QQSel_X = C_Sel & NLT & X
End Function
Function QQSelDis_X$(X)
QQSelDis_X = C_SelDis & NLT & X
End Function

Function QQSel_FF$(FF)
QQSel_FF = C_Sel & NLT & FF_JnComma(FF)
End Function
Function QQSelDis_FF$(FF)
QQSelDis_FF = C_SelDis & NLT & FF_JnComma(FF)
End Function

Function QQSet_FF_Ey$(FF, Ey$())
Dim Fny$(): Fny = Ssl_Sy(FF)
Ay_XAss_SamSz Fny, Ey, CSub, "FF", "Ey"
Dim AFny$()
    AFny = Ay_XAlign_L(Fny)
    AFny = Ay_XAdd_Sfx(AFny, " = ")
Dim W%
    W = VblAy_Wdt(Ey)
Dim Ident%
    W = Ay_Wdt(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        Push Ay, VblAlign(Ey(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
        Push Ay1, VblAlign(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
QQSet_FF_Ey = Vbl
End Function

Private Sub Z_SetFld()
Dim Fny$(), ExprVblAy$()
Fny = Ssl_Sy("a b c d")
Push ExprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "2sdfkl|lskdfjdf| sdf"
Push ExprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "4sf| sdf"
    Act = SetFld(Fny, ExprVblAy)
Debug.Print XRpl_VBar(Act)
End Sub
Function SetFld$(Fny$(), Ay)

End Function
Private Function Q$(A)
Q = XQuote_SqBkt_IfNeed(A)
End Function
Function Upd$(T)
Upd = C_Upd & NLT & Q(T)
End Function
Function DbtSkVV_BExpr$(A As Database, T, SkVV)
DbtSkVV_BExpr = FnyVV_BExpr(Dbt_SkFny(A, T), SkVV)
End Function
Function FnyVV_BExpr$(Fny$(), VV)
Dim O$(), J%, F, Vy()
Vy = CvVV(VV)
For Each F In Fny
    PushI O, XQuote_SqBkt_IfNeed(F) & "=" & Val_XQuote_Sql(Vy(J))
    J = J + 1
Next
FnyVV_BExpr = JnAnd(O)
End Function
Function Val_XQuote_Sql$(A)
Select Case True
Case IsStr(A): Val_XQuote_Sql = XQuote_Sng(A)
Case IsDte(A): Val_XQuote_Sql = XQuote_Dte(A)
Case Else: Val_XQuote_Sql = A
End Select
End Function

Function Whfv(F, V) ' Ssk is single-Sk-value
Whfv = CWh & Q(F) & "=" & QV(V)
End Function

Function WhK$(K&, T)
WhK = Whfv(T & "Id", K)
End Function

Function WhBet$(F, FmV, ToV)
WhBet = CWh & Q(F) & QV(FmV) & CAnd & QV(ToV)
End Function

Private Function QV$(V)
QV = Var_XQuote_Sql(V)
End Function
Private Property Get CAnd$()
CAnd = " " & C_And & " "
End Property
Private Property Get CWh$()
CWh = NLT & C_Wh & NLTT
End Property
Function Wh$(BExpr$)
If BExpr = "" Then Exit Function
Wh = CWh & BExpr
End Function
Function WhFldInAy$(F, InAy)
WhFldInAy = CWh & FldInAy(F, InAy)
End Function
Private Sub Z_WhFldInAy()
Dim Fny$(), Ay()
Fny = Ssl_Sy("A B C")
Ay = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = WhFldInAy(Fny, Ay)
    C
    Return
End Sub

Private Function FnyEqVy_BExpr$(Fny$(), EqVy)

End Function

Function QQWh_FnyEqVy$(Fny$(), EqVy)
QQWh_FnyEqVy = CWh & FnyEqVy_BExpr(Fny, EqVy)
End Function

Function QQInBExpr_F_Ay$(Ay, FldNm$, Optional WithQuote As Boolean)
Const C$ = "[?] in (?)"
Dim B$
    If WithQuote Then
        B = JnComma(AyQuoteSng(Ay))
    Else
        B = JnComma(Ay)
    End If
QQInBExpr_F_Ay = QQ_Fmt(C, FldNm, B)
End Function

Function QQWh_BExpr$(A$)
If A = "" Then Exit Function
QQWh_BExpr = C_NL & "Where" & NLT & A
End Function

Private Sub Z_FnyVy_SetSqpFmt()
Dim Fny$(), Vy()
Ept = XRpl_VBar("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = Lin_TermAy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = FnyVy_SetSqp(Fny, Vy)
    C
    Return
End Sub

Private Sub Z_WhFldInAySqpAy()

End Sub

Function VblAy_AlignAsLines$(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAy, Optional Sep$ = ",")
VblAy_AlignAsLines = JnVBar(VblAy_AlignAsLy(ExprVblAy, Pfx, IdentOpt, SfxAy, Sep))
End Function

Function VblAy_AlignAsLy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAyOpt, Optional Sep$ = ",") As String()
Dim NoSfxAy As Boolean
Dim SfxWdt%
Dim SfxAy$()
    NoSfxAy = IsEmp(SfxAy)
    If Not NoSfxAy Then
        Ass IsSy(SfxAyOpt)
        SfxAy = Ay_XAlign_L(SfxAyOpt)
        Dim U%, J%: U = UB(SfxAy)
        For J = 0 To U
            If J <> U Then
                SfxAy(J) = SfxAy(J) & Sep
            End If
        Next
    End If
Ass IsVblAy(ExprVblAy)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$
U = UB(ExprVblAy)
Dim W%
    W = VblAy_Wdt(ExprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If NoSfxAy Then
        If J = U Then S = "" Else S = Sep
    Else
        If J = U Then S = SfxAy(J) Else S = SfxAy(J) & Sep
    End If
    Push O, VblAlign(ExprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
VblAy_AlignAsLy = O
End Function

Function QQSel_Fny_EDic$(Fny$(), EDic As Dictionary)
Stop '
QQSel_Fny_EDic = "" 'QSel_Fny_Ey(Fny, DicKySy(EDic, Fny))
End Function

Function QQSel_Fny_Ey$(Fny$(), Ey$())
QQSel_Fny_Ey = QQSel_X(FnyEy_X(Fny, Ey))
End Function

Private Function FnyEy_X(Fny$(), Ey$())
If IsSqlFmt Then
    FnyEy_X = FnyEy_AsLines(Fny, Ey)
Else
    FnyEy_X = FnyEy_AsLin(Fny, Ey)
End If
End Function

Private Function FnyEy_AsLin$(Fny$(), Ey$())
'AyAB_IsSamSz_XAss Fny, Ey, CSub, "Fny Ey"
Dim O$()
    Dim F, J%
    For Each F In Fny
        If F = Ey(J) Or Ey(J) = "" Then
            PushI O, XQuote_SqBkt_IfNeed(F)
        Else
            PushI O, XQuote_SqBkt_IfNeed(Ey(J)) & " As " & XQuote_SqBkt_IfNeed(F)
        End If
        J = J + 1
    Next
FnyEy_AsLin = JnCommaSpc(O)
End Function

Private Function FnyEy_AsLines(Fny$(), Ey$())
Dim N$(), E$()
N = Ay_XAlign_L(Fny)
E = Ay_XAlign_L(Ay_XQuote_SqBkt_IfNeed(Ey))
Dim O$(), J%
For J = 0 To UB(Fny)
    If Fny(J) = Ey(J) Or Ey(J) = "" Then
        Push O, QQ_Fmt("     ?    ?", Space(Len(E(J))), N(J))
    Else
        Push O, QQ_Fmt("     ? As ?", E(J), N(J))
    End If
Next
  
FnyEy_AsLines = Jn(O, "," & NLTT)
End Function

Private Function FF_Fy(FF) As String()
FF_Fy = Ay_XQuote_SqBkt_IfNeed(FF_Fny(FF))
End Function
Function FFExprAy_AsX$(FF, Ey$())
If IsSqlFmt Then
    FFExprAy_AsX = FFExprAy_AsLines(FF, Ey)
Else
    FFExprAy_AsX = FFExprAy_AsLin(FF, Ey)
End If
End Function
Function FFExprAy_AsLines$(FF, Ey$())
FFExprAy_AsLines = FnyEy_AsLines(FF_Fy(FF), Ey)
End Function

Function FFExprAy_AsLin$(FF, Ey$())
FFExprAy_AsLin = FnyEy_AsLin(FF_Fy(FF), Ey)
End Function

Function Fny_FF$(A$())
Fny_FF = JnSpc(Ay_XQuote_SqBkt_IfNeed(A))
End Function

Function FF_Fny(FF) As String()
Select Case True
Case IsStr(FF): FF_Fny = Lin_TermAy(FF)
Case IsSy(FF): FF_Fny = FF
Case Else: XThw CSub, "Given FF must be Str or Sy", "TypeName FF", TypeName(FF), FF
End Select
End Function

Private Function FFExprDic_AsLines$(FF, E As Dictionary)
Dim Fny$(): Fny = FF_Fny(FF)
Dim Ey$(): Ey = DicKySy(E, Fny)
FFExprDic_AsLines = FFExprAy_AsLines(Fny, Ey)
End Function
Private Sub Z_QQWh_FnyEqVy()

End Sub

Private Sub Z()
Z_Gp
Z_SetFld
Z_WhFldInAy
Z_WhFldInAySqpAy
Z_QQWh_FnyEqVy
Sql_Shared:
End Sub

Private Function Var_QuoteChr$(A)
Dim O$
Select Case True
Case IsStr(A): O = "'"
Case IsDate(A): O = "#"
Case IsEmpty(A), IsNull(A), IsNothing(A): Stop
Var_QuoteChr = O
End Select
End Function
Function Vy_XQuote_Sql(Vy) As String()
Dim V
For Each V In Vy
    PushI Vy_XQuote_Sql, Var_XQuote_Sql(V)
Next
End Function

Function Var_XQuote_Sql$(A)
Dim Q$
Q = Var_QuoteChr(A)
Var_XQuote_Sql = Q & A & Q
End Function


