Attribute VB_Name = "MVb_Ay"
Option Compare Binary
Option Explicit

Sub AyAsg(A, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To Min(UB(Av), UB(A))
    Asg A(J), OAp(J)
Next
End Sub

Function AyAsgAy(A, OIntoAy)
If TypeName(A) = TypeName(OIntoAy) Then
    OIntoAy = A
    AyAsgAy = OIntoAy
    Exit Function
End If
If Sz(A) = 0 Then
    Erase OIntoAy
    AyAsgAy = OIntoAy
    Exit Function
End If
Dim U&
    U = UB(A)
ReDim OIntoAy(U)
Dim I, J&
For Each I In A
    Asg I, OIntoAy(J)
    J = J + 1
Next
AyAsgAy = OIntoAy
End Function

Sub AyAsgT1AyRestAy(A, OT1Ay$(), ORestAy$())
Dim U&, J&
U = UB(A)
If U = -1 Then
    Erase OT1Ay, ORestAy
    Exit Sub
End If
ReDim OT1Ay(U)
ReDim ORestAy(U)
For J = 0 To U
    BrkAsg A(J), " ", OT1Ay(J), ORestAy(J)
Next
End Sub

Sub Ay_XBrw_XHalt(A, Optional Fnn$)
If Sz(A) = 0 Then Exit Sub
Ay_XBrw A, Fnn
XHalt
End Sub
Function Ay_XBrw(A, Optional Fnn$)
Dim T$
T = TmpFt("Ay_XBrw", Fnn)
Ay_XWrt A, T
Ft_XBrw T
End Function

Function Ay_XCln(A)
If Not IsArray(A) Then XThw CSub, "Given [A] is not an array, but [TypeName]", "TypeName", TypeName(A)
Dim O
O = A
Erase O
Ay_XCln = O
End Function

Function Ay_XChk_Dup(A, QMsg$) As String()
Dim Dup
Dup = Ay_XWh_Dup(A)
If Sz(Dup) = 0 Then Exit Function
Push Ay_XChk_Dup, QQ_Fmt(QMsg, JnSpc(Dup))
End Function

Function AyDupT1(A) As String()
AyDupT1 = Ay_XWh_Dup(Ay_XTak_T1(A))
End Function

Function AyEmpChk(A, Msg$) As String()
If Sz(A) = 0 Then AyEmpChk = Ap_Sy(Msg)
End Function

Function Ay_IsEqChk(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, QQ_Fmt("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
If Sz(Ay1) = 0 Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, QQ_Fmt("[?]-th Ele is diff: ?[?]<>?[?]", Ay1Nm, Ay2Nm, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
If IsEmp(O1) Then Exit Function
Dim O2$()
    Push O2, QQ_Fmt("Array [?] and [?] both having size[?] have differnt element(s):", Ay1Nm, Ay2Nm, Sz(Ay1))
    If ReachLimit Then
        Push O2, QQ_Fmt("At least [?] differences:", Sz(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, QQ_Fmt("Ay-[?]:", Ay1Nm)
PushAy O, AyQuote(Ay1, "[]")
Push O, QQ_Fmt("Ay-[?]:", Ay2Nm)
PushAy O, AyQuote(Ay2, "[]")
Ay_IsEqChk = O
End Function

Function AyOfAy_Ay(AyOfAy)
If Sz(AyOfAy) = 0 Then Exit Function
Dim O
O = Ay_XCln(AyOfAy(0))
Dim X
For Each X In AyOfAy
    PushAy O, X
Next
AyOfAy_Ay = O
End Function
Private Sub Z_AyFlat()
Dim AyOfAy()
AyOfAy = Array(Ssl_Sy("a b c d"), Ssl_Sy("a b c"))
Ept = Ssl_Sy("a b c d a b c")
GoSub Tst
Exit Sub
Tst:
    Act = AyFlat(AyOfAy)
    C
    Return
End Sub

Function AyFlat(AyOfAy)
AyFlat = AyOfAy_Ay(AyOfAy)
End Function

Function AyItmCnt%(A, M)
If Sz(A) = 0 Then Exit Function
Dim O%, X
For Each X In AyNz(A)
    If X = M Then O = O + 1
Next
AyItmCnt = O
End Function

Function Ay_XWh_LasN(A, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(A)
If U < N Then Ay_XWh_LasN = A: Exit Function
O = A
Fm = U - N + 1
NewU = N - 1
For J = Fm To U
    Asg O(J), O(I)
    I = I + 1
Next
ReDim Preserve O(NewU)
Ay_XWh_LasN = O
End Function

Function Ay_LasEle(A)
Asg A(UB(A)), Ay_LasEle
End Function

Function AyLines$(A)
AyLines = JnCrLf(A)
End Function

Function AyMid(A, Fm, Optional L = 0)
AyMid = Ay_XCln(A)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(A)
    Case Else:  E = Min(UB(A), L + Fm - 1)
    End Select
For J = Fm To E
    Push AyMid, A(J)
Next
End Function


Function AyNPfxStar%(A)
Dim O%, X
For Each X In AyNz(A)
    If XTak_FstChr(X) = "*" Then AyNPfxStar = O: Exit Function
    O = O + 1
Next
End Function
Function AyNxtNm$(A, Nm$, Optional MaxN% = 99)
If Not Ay_XHas(A, Nm) Then AyNxtNm = Nm: Exit Function
Dim J%, O$
For J = 1 To MaxN
    O = Nm & Format(J, "00")
    If Not Ay_XHas(A, O) Then AyNxtNm = O: Exit Function
Next
Stop
End Function
Function ItrNz(A)
If IsNothing(A) Then Set ItrNz = New Collection Else Set ItrNz = A
End Function
Function AyNz(A)
If Sz(A) = 0 Then Set AyNz = New Collection Else AyNz = A
End Function

Sub AyPushMsgAv(A, Msg$, Av())
PushAy A, MsgAv_Ly(Msg, Av)
End Sub


Function AyRTrim(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), I
For Each I In A
    Push O, RTrim(I)
Next
AyRTrim = O
End Function

Function Ay_XReSz(A, SzAy) ' return Ay as [A] with sz as [SzAy]
If Sz(SzAy) = 0 Then Ay_XReSz = A: Exit Function
Dim O
O = A
ReDim O(UB(SzAy))
Ay_XReSz = O
End Function

Function AyReverseI(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = A(U - J)
Next
AyReverseI = O
End Function

Function AyReverseObj(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    Set O(J) = A(U - J)
Next
AyReverseObj = O
End Function

Function Ay_XRpl_MidSeg_FT_IX(A, B As FTIx, AySeg)
Ay_XRpl_MidSeg_FT_IX = Ay_XRpl_MidSeg(A, B.FmIx, B.ToIx, AySeg)
End Function

Function Ay_XRpl_MidSeg(A, FmIx&, ToIx&, ByAy)
Dim M As AyABC
    M = Ay_XBrk_ABC(A, FmIx, ToIx)
Ay_XRpl_MidSeg = M.A
    PushAy Ay_XRpl_MidSeg, ByAy
    PushAy Ay_XRpl_MidSeg, M.C
End Function

Function Ay_XRpl_Star_InEach_Ele(A$(), By) As String()
Dim X
For Each X In AyNz(A)
    PushI Ay_XRpl_Star_InEach_Ele, Replace(X, By, "*")
Next
End Function

Function Ay_XRpl_T1(A$(), T1$) As String()
Ay_XRpl_T1 = Ay_XAdd_Pfx(Ay_XRmv_T1(A), T1 & " ")
End Function

Function Ay_SampLin$(A)
Dim S$, U&
U = UB(A)
If U >= 0 Then
    Select Case True
    Case IsPrim(A(0)): S = "[" & A(0) & "]"
    Case IsObject(A(0)), IsArray(A(0)): S = "[*Ty:" & TypeName(A(0)) & "]"
    Case Else: Stop
    End Select
End If
Ay_SampLin = "*Ay:[" & U & "]" & S
End Function

Function Itm_IsSel_ByAy(Itm, Ay) As Boolean
If Sz(Ay) = 0 Then Itm_IsSel_ByAy = True: Exit Function
Itm_IsSel_ByAy = Ay_XHas(Ay, Itm)
End Function

Function Ay_SeqCntDic(A) As Dictionary 'The return dic of key=AyEle pointing to Long(1) with Itm0 as Seq# and Itm1 as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In AyNz(A)
    If O.Exists(X) Then
        L = O(X)
        L(1) = L(1) + 1
        O(X) = L
    Else
        ReDim L(1)
        L(0) = S
        L(1) = 1
        O.Add X, L
    End If
Next
Set Ay_SeqCntDic = O
End Function
Function AySqH(A) As Variant()
Dim N&: N = Sz(A)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To 1, 1 To N)
For Each V In A
    J = J + 1
    O(1, J) = V
Next
AySqH = O
End Function

Function AySqV(A) As Variant()
Dim N&: N = Sz(A)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To N, 1 To 1)
For Each V In A
    J = J + 1
    O(J, 1) = V
Next
AySqV = O
End Function

Function AyT1Chd(A, T1) As String()
AyT1Chd = Ay_XRmv_T1(Ay_XWh_T1(A, T1))
End Function

Function AyIdent(A, Optional Ident% = 4) As String()
Dim I, S$
S = Space(Ident)
For Each I In AyNz(A)
    PushI AyIdent, S & I
Next
End Function
Function AyTrim(A) As String()
Dim X
For Each X In AyNz(A)
    Push AyTrim, Trim(X)
Next
End Function

Function Ay_Wdt%(A)
Dim O%, J&
For J = 0 To UB(A)
    O = Max(O, Len(A(J)))
Next
Ay_Wdt = O
End Function

Function AyWrpPad(A, W%) As String() ' Each Itm of Ay[A] is padded to line with Ay_Wdt(A).  return all padded lines as String()
Dim O$(), X, I%
ReDim O(0)
For Each X In AyNz(A)
    If Len(O(I)) + Len(X) < W Then
        O(I) = O(I) & X
    Else
        PushI O, X
        I = I + 1
    End If
Next
AyWrpPad = O
End Function

Sub Ay_XWrt(A, Ft)
Str_XWrt JnCrLf(A), Ft
End Sub

Function Ay_XZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O(), J&
ReSz O, U
For J = 0 To U
    If U1 >= J Then
        If U2 >= J Then
            O(J) = Array(A1(J), A2(J))
        Else
            O(J) = Array(A1(J), Empty)
        End If
    Else
        If U2 >= J Then
            O(J) = Array(, A2(J))
        Else
            Stop
        End If
    End If
Next
Ay_XZip = O
End Function

Function Ay_XZip_Ap(A, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
Dim UCol%
    UCol = UB(Av)

Dim URow1&
    URow1 = UB(A)

Dim URow&
Dim URowAy&()
    Dim J%, IURow%
    URow = URow1
    For J = 0 To UB(Av)
        IURow = UB(Av(J))
        Push URowAy, IURow
        If IURow > URow Then URow = IURow
    Next

Dim ODry()
    Dim Dr()
    ReSz ODry, URow
    Dim I%
    For J = 0 To URow
        Erase Dr
        If URow1 >= J Then
            Push Dr, A(J)
        Else
            Push Dr, Empty
        End If
        For I = 0 To UB(Av)
            If URowAy(I) >= J Then
                Push Dr, Av(I)(J)
            Else
                Push Dr, Empty
            End If
        Next
        ODry(J) = Dr
    Next
Ay_XZip_Ap = ODry
End Function


Function ItmAddAy(Itm, Ay)
ItmAddAy = AyIns(Ay, Itm)
End Function

Function SSAy_Sy(SSAy) As String()
Dim SS
For Each SS In AyNz(SSAy)
    PushIAy SSAy_Sy, Ssl_Sy(SS)
Next
End Function

Function VyDr(A(), FF, Fny$()) As Variant()
Dim IxAy&(), U%
    IxAy = Ay_IxAy(Fny, Ssl_Sy(FF))
    U = AyMax(IxAy)
    GoSub X_ChkIxAy
Dim O(), J%, Ix
ReDim O(U)
For Each Ix In IxAy
    O(Ix) = A(J)
    J = J + 1
Next
VyDr = O
Exit Function
X_ChkIxAy:
    For Each Ix In IxAy
        If Ix <= -1 Then Stop
    Next
    Return
End Function

Private Sub ZZZ_Ay_XBrk_ABC()
Dim A(): A = Array(1, 2, 3, 4)
Dim Act As AyABC: Act = Ay_XBrk_ABC(A, 1, 2)
Ass Ay_IsEq(Act.A, Array(1))
Ass Ay_IsEq(Act.B, Array(2, 3))
Ass Ay_IsEq(Act.C, Array(4))
End Sub

Private Sub ZZ_AyAsgAp()
Dim O%, A$
'AyAsgAp Array(234, "abc"), O, A
Ass O = 234
Ass A = "abc"
End Sub

Private Sub ZZ_Ay_IsEqChk()
Ay_XDmp Ay_IsEqChk(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub



Private Sub ZZ_AyMax()
Dim A()
Dim Act
Act = AyMax(A)
Stop
End Sub

Private Sub ZZ_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
Ay_IsEq_XAss Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
Ay_IsEq_XAss Exp, Act
End Sub

Private Sub ZZ_Ay_XExl_EmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = Ay_XExl_EmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub ZZ_Ay_Sy()
Dim Act$(): Act = Ay_Sy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub ZZ_AyTrim()
Ay_XDmp AyTrim(Array(1, 2, 3, "  a"))
End Sub


Private Sub Z_Ay_XChk_Dup()
Dim Ay
Ay = Array("1", "1", "2")
Ept = Ap_Sy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XChk_Dup(Ay, "This item[?] is duplicated")
    C
    Return
End Sub

Private Sub Z_Ay_IsEqChk()
Ay_XDmp Ay_IsEqChk(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub Z_Ay_XBrk_ABC_FTIx()
Dim A(): A = Array(1, 2, 3, 4)
Dim M As FTIx: M = New_FTIx(1, 2)
Dim Act As AyABC: Act = Ay_XBrk_ABC_FTIx(A, M)
Ass Ay_IsEq(Act.A, Array(1))
Ass Ay_IsEq(Act.B, Array(2, 3))
Ass Ay_IsEq(Act.C, Array(4))
End Sub

Private Sub Z_Ay_XHasDupEle()
Ass Ay_XHasDupEle(Array(1, 2, 3, 4)) = False
Ass Ay_XHasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Private Sub Z_AyIns()
Dim A, M, At&
'
A = Array(1, 2, 3)
M = "X"
Ept = Array("X", 1, 2, 3)
GoSub Tst
'
Exit Sub
Tst:
    Act = AyIns(A, M, At)
    C
Return
End Sub

Private Sub Z_AyInsAy()
Dim Act, Exp, A(), B(), At&
A = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAy(A, B, At)
Ass Ay_IsEq(Act, Exp)
End Sub

Private Sub Z_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
Ay_XAss_Eq Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
Ay_XAss_Eq Exp, Act
End Sub

Private Sub Z_Ay_Sy()
Dim Act$(): Act = Ay_Sy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub Z_AyTrim()
Ay_XDmp AyTrim(Ap_Sy(1, 2, 3, "  a"))
End Sub

Private Sub Z_Aydic_to_KeyCntMulItmCol_Dry()
Dim A As New Dictionary, Act()
A.Add "A", Array(1, 2, 3)
A.Add "B", Array(2, 3, 4)
A.Add "C", Array()
A.Add "D", Array("X")
Act = Aydic_to_KeyCntMulItmColDry(A)
Ass Sz(Act) = 4
Ass Ay_IsEq(Act(0), Array("A", 3, 1, 2, 3))
Ass Ay_IsEq(Act(1), Array("B", 3, 2, 3, 4))
Ass Ay_IsEq(Act(2), Array("C", 0))
Ass Ay_IsEq(Act(3), Array("D", 1, "X"))
End Sub
Private Sub Z_VyDr()
Dim Fny$(), FF, Vy()
Fny = Ssl_Sy("A B C D E f")
FF = "C E"
Vy = Array(1, 2)
Ept = Array(Empty, Empty, 1, Empty, 2)
GoSub Tst
Exit Sub
Tst:
    Act = VyDr(Vy, FF, Fny)
    C
    Return
End Sub

Function CvAy(A) As Variant()
CvAy = A
End Function

Function CvAyITM(Itm_or_Ay) As Variant()
If IsArray(Itm_or_Ay) Then
    CvAyITM = Itm_or_Ay
Else
    CvAyITM = Array(Itm_or_Ay)
End If
End Function



Private Sub Z()
Z_Aydic_to_KeyCntMulItmCol_Dry
Z_Ay_XChk_Dup
Z_AyFlat
Z_Ay_XBrk_ABC_FTIx
Z_Ay_XHasDupEle
Z_AyIns
Z_AyInsAy
Z_Ay_IsEqChk
Z_AyMinus
Z_Ay_Sy
Z_AyTrim
Z_VyDr
MVb_Ay:
End Sub
