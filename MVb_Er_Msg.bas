Attribute VB_Name = "MVb_Er_Msg"
Option Compare Binary
Option Explicit

Function FunMsgAp_Lin$(A$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAp_Lin = FunMsgAp_Lin(A$, Msg, Av)
End Function

Sub FunMsgAp_XDmpLin(A$, MsgSV$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAv_XDmp_Lin A, MsgSV, Av
End Sub

Sub FunMsgAp_XDmp_Ly(A, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgAv_Ly(A, Msg, Av)
End Sub

Function FunMsgAv_Ly(A, Msg$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(Msg)
C = NyAv_Ly(CvSy(Ay_XAdd_(Ap_Sy("Fun"), Msg_Ny(Msg))), CvAy(Ay_XAdd_(Array(A), Av)))
FunMsgAv_Ly = Ay_XAdd_(B, C)
End Function

Sub FunMsgAv_XBrw(A, Msg$, Av())
Ay_XBrw FunMsgAv_Ly(A, Msg, Av)
End Sub

Sub XDmp_Lin(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAv_XDmp_Lin Fun, Msg, Av
End Sub

Sub XDmp_Lin_AV(Fun$, Msg$, Av())
FunMsgAv_XDmp_Lin Fun, Msg, Av
End Sub

Sub FunMsgAv_XDmp_Lin(A$, Msg$, Av())
D FunMsgAp_Lin(A, Msg, Av)
End Sub

Sub FunMsgAv_XDmp_Ly(A$, Msg$, Av())
D FunMsgAv_Ly(A, Msg, Av)
End Sub

Private Function FunDotMsg_Fmt(Fun$, DotMsg$) As String()
Dim O$()
O = DotMsg_Fmt(DotMsg)
If Sz(O) = 0 Then
    PushI O, "@" & Fun
Else
    O(0) = O(0) & "  " & "@" & Fun
End If
FunDotMsg_Fmt = O
End Function

Function FunMsgNyAp_Ly(Fun$, Msg$, Ny0, ParamArray Ap()) As String()
Dim Av(): Av = Ap
FunMsgNyAp_Ly = FunMsgNyAv_Ly(Fun, Msg, Ny0, Av)
End Function

Sub FunMsgNyAp_XDmp(A$, Msg$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgNyAv_Ly(A, Msg, Ny0, Av)
End Sub

Function FunMsgNyAv_Lin$(Fun$, DotMsg$, Ny0, Av())
FunMsgNyAv_Lin = Ap_JnVBarSpc(Fun, DotMsg, NyAv_Lin(Ny0, Av))
End Function

Function FunMsgNyAv_Ly(Fun$, DotMsg$, Ny0, Av()) As String()
FunMsgNyAv_Ly = Ay_XAdd_(FunDotMsg_Fmt(Fun, DotMsg), NyAv_Ly_IDENT(Ny0, Av, 4))
End Function

Function LvlSep$(Lvl%)
Select Case Lvl
Case 0: LvlSep = "."
Case 1: LvlSep = "-"
Case 2: LvlSep = "+"
Case 3: LvlSep = "="
Case 4: LvlSep = "*"
Case Else: LvlSep = Lvl
End Select
End Function

Function MacroStrAvLy(A$, Av()) As String()
MacroStrAvLy = NyAv_Ly(MacroNy(A, OpnBkt:="["), Av)
End Function

Sub Msg(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgNyAv_Ly(Fun, Msg, MacroNy(Msg), Av)
End Sub

Sub MsgApBrw(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_XBrw Msg, Av
End Sub

Sub MsgAp_XDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
Ay_XDmp MsgAv_Ly(A, Av)
End Sub

Function MsgAp_Lin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAp_Lin = MsgAv_Lin(A, Av)
End Function

Function MsgAp_Ly(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgAp_Ly = MsgAv_Ly(A, Av)
End Function

Sub MsgAp_XBrw(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_XBrw A, Av
End Sub

Sub MsgAp_XBrwStop(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_XBrw A, Av
Stop
End Sub

Sub MsgAp_XDmpLin(A$, ParamArray Ap())
Dim Av(): Av = Ap
D MsgAv_Lin(A, Av)
End Sub

Sub MsgAp_XDmp_Ly(A$, ParamArray Ap())
Dim Av(): Av = Ap
Ay_XDmp MsgAv_Ly(A, Av)
End Sub

Function MsgAv_Lin$(A$, Av())
Dim B$(), C$
C = NyAv_Lin(Msg_Ny(A), Av)
MsgAv_Lin = XEns_SfxDot(A) & C
End Function

Function MsgAv_Ly(A$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(A)
C = AyTab(NyAv_Ly(Msg_Ny(A), Av))
MsgAv_Ly = Ay_XAdd_(B, C)
End Function

Sub MsgAv_XBrw(A$, Av())
Ay_XBrw MsgAv_Ly(A, Av)
End Sub

Function DotMsg_Fmt(A$) As String()
Dim A1$, A2$
    Brk1Asg A, ".", A1, A2
PushI DotMsg_Fmt, A1 & ".  "
If A2 = "" Then Exit Function
PushIAy DotMsg_Fmt, Ay_XAdd_Pfx(Lines_XWrap(A2), "==> ")
End Function

Function MsgNyAp_Ly(Msg$, Ny0, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgNyAp_Ly = MsgNyAv_Ly(Msg, Ny0, Av)
End Function
Sub XDmp_Lin_Stop(Fun$, Msg$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgNyAv_XDmp_Lin Fun, Msg, Ny0, Av
XShw_Dbg
Stop
End Sub
Sub MsgNyAp_XDmp_Lin_Stop(Msg$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
D MsgNyAv_Ly(Msg, Ny0, Av)
XShw_Dbg
Stop
End Sub

Function MsgNyAv_Ly(MsgDot$, Ny0, Av()) As String()
MsgNyAv_Ly = Ay_XAdd_(DotMsg_Fmt(MsgDot), NyAv_Ly_IDENT(Ny0, Av, 4))
End Function

Sub MsgNyAv_XDmp_Ly(Fun$, Msg$, FF, Av())
D FunMsgNyAv_Ly(Fun, Msg, FF, Av)
End Sub

Sub FunMsgNyAv_XDmp_Lin(Fun$, Msg$, FF, Av())
D FunMsgNyAv_Lin(Fun, Msg, FF, Av)
End Sub

Sub FunMsgNyAp_Lin(Fun$, Msg$, FF, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgNyAv_Lin(Fun, Msg, FF, Av)
End Sub

Sub MsgObj_Prp(Fun$, Msg$, Obj, PrpNy0)
MsgNyAv_XDmp_Ly Fun, Msg, PrpNy0, Obj_PrpAy(Obj, PrpNy0)
End Sub

Sub MsgAp_XHalt(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D MsgNyAv_Ly(Msg, MacroNy(Msg), Av)
End Sub

Function Msg_Ny(A) As String()
Dim O$(), P%, J%
O = Split(A, "[")
Ay_XShf_ O
For J = 0 To UB(O)
    P = InStr(O(J), "]")
    O(J) = "[" & Left(O(J), P)
Next
Msg_Ny = O
End Function

Function NmV_Ly(Nm$, V) As String()
Dim Ly$(): Ly = Var_Ly(V)
If IsDic(V) Then
    Stop
End If
Dim J%, S$
If Sz(Ly) = 0 Then
    PushI NmV_Ly, Nm & ": "
Else
    PushI NmV_Ly, Nm & ": " & Ly(0)
End If
S = Space(Len(Nm) + 2)
For J = 1 To UB(Ly)
    PushI NmV_Ly, S & Ly(J)
Next
End Function

Function NmV_Str$(Nm$, V)
NmV_Str = Nm & "=[" & Var_Str(V) & "]"
End Function

Function NyAp_Ly(Ny0, ParamArray Ap()) As String()
Dim Av(): Av = Ap
NyAp_Ly = NyAv_Ly(Ny0, Av)
End Function

Sub NyAp_XDmp(Ny0, ParamArray Ap())
Dim Av(): Av = Ap
D NyAv_Ly(Ny0, Av)
End Sub

Function NyAv_Lin$(Ny0, Av())
Dim U&, Ny$()
Ny = CvNy(Ny0)
U = UB(Ny)
If U = -1 Then Exit Function
Dim O$(), J%
For J = 0 To U
    Push O, NmV_Str(Ny(J), Av(J))
Next
NyAv_Lin = JnVBarSpc(O)
End Function

Function NyAv_Ly(Ny0, Av()) As String()
Dim J%, Ny$()
Ny = CvNy(Ny0)
Ny = Ay_XAlign_L(Ny)
AyAB_XSet_SamMax Ny, Av
For J = 0 To UB(Ny)
    PushIAy NyAv_Ly, NmV_Ly(Ny(J), Av(J))
Next
End Function

Function NyAv_Ly_IDENT(Ny0, Av(), Optional Ident% = 4) As String()
NyAv_Ly_IDENT = AyIdent(NyAv_Ly(Ny0, Av), Ident)
End Function

Function NyAv_Scl$(A$(), Av())
Dim O$(), J%, X, Y
X = A
Y = Av
AyAB_XSet_SamMax X, Y
For J = 0 To UB(X)
    Push O, XRmv_SqBkt(X(J)) & "=" & Var_Str(Y(J))
Next
NyAv_Scl = JnSemiColon(O)
End Function

Function Var_Lines$(A, Optional Lvl%)
Dim T$, S$, W%, I, O$(), Sep$
Select Case True
Case IsDic(A): Var_Lines = JnCrLf(Dic_Fmt(CvDic(A)))
Case IsPrim(A): Var_Lines = A
Case IsLinesAy(A): Var_Lines = LinesAy_Lines(CvSy(A))
Case IsSy(A): Var_Lines = JnCrLf(A)
Case IsNothing(A): Var_Lines = "#Nothing"
Case IsEmpty(A): Var_Lines = "#Empty"
Case IsMissing(A): Var_Lines = "#Missing"
Case IsObject(A): Var_Lines = "#Obj(" & TypeName(A) & ")"
Case IsArray(A)
    If Sz(A) = 0 Then Exit Function
    For Each I In A
        PushI O, Var_Lines(I, Lvl + 1)
    Next
    If Lvl > 0 Then
        W = LinesAy_Wdt(O)
        Sep = LvlSep(Lvl)
        PushI O, StrDup(Sep, W)
    End If
    Var_Lines = JnCrLf(O)
Case Else
End Select
End Function

Function Var_Ly(A) As String()
Var_Ly = SplitCrLf(Var_Lines(A))
End Function

Function Var_Str$(V)
Select Case True
Case IsPrim(V):    Var_Str = V
Case IsArray(V):   Var_Str = AyLines(V)
Case IsNothing(V): Var_Str = "*Nothing"
Case IsObject(V):  Var_Str = "*Type[" & TypeName(V) & "]"
Case IsEmpty(V):   Var_Str = "*Empty"
Case IsMissing(V): Var_Str = "*Missing"
Case Else: Stop
End Select
End Function

Sub XDmp_Ly(Fun$, Msg$, FF, ParamArray Ap())
Dim Av(): Av = Ap
MsgNyAv_XDmp_Ly Fun, Msg, FF, Av
End Sub

Private Sub Z_MsgObj_Prp()
Dim Fun$, Msg$, Obj, PrpNy0
Fun = "XXX"
Msg = "MsgABC"
Set Obj = New DAO.Field
PrpNy0 = "Name Type Size"
GoSub Tst
Exit Sub
Tst:
    MsgObj_Prp Fun, Msg, Obj, PrpNy0
End Sub

Private Sub ZZ()
Dim A$
Dim B()
Dim C
Dim D%
Dim F$()
Dim XX
FunMsgAp_Lin A, A, B
FunMsgAp_XDmpLin A, A, B
FunMsgAp_XDmp_Ly C, A, B
FunMsgAv_Ly C, A, B
FunMsgAv_XBrw C, A, B
FunMsgNyAp_Ly A, A, C, B
FunMsgNyAp_XDmp A, A, C, B
FunMsgNyAv_Lin A, A, C, B
FunMsgNyAv_Ly A, A, C, B
LvlSep D
MacroStrAvLy A, B
Msg A, A, B
MsgApBrw A, B
MsgAp_XDmp A, B
MsgAp_Lin A, B
MsgAp_Lin A, B
MsgAp_Ly A, B
MsgAp_XBrw A, B
MsgAp_XBrwStop A, B
MsgAp_XDmpLin A, B
MsgAp_XDmp_Ly A, B
MsgAv_Lin A, B
MsgAv_Ly A, B
MsgAv_XBrw A, B
DotMsg_Fmt A
MsgNyAp_Ly A, C, B
MsgNyAp_XDmp_Lin_Stop A, C, B
MsgNyAv_Ly A, C, B
MsgNyAv_XDmp_Ly A, A, C, B
End Sub

Private Sub Z()
Z_MsgObj_Prp
End Sub
