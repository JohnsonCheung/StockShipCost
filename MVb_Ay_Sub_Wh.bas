Attribute VB_Name = "MVb_Ay_Sub_Wh"
Option Compare Binary
Option Explicit

Sub AyDupAss(A, Fun$, Optional IgnCas As Boolean)
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = Ay_XWh_Dup(A, IgnCas)
If Sz(Dup) = 0 Then Exit Sub
XThw Fun, "There are dup in array", "Dup Ay", Dup, A
End Sub

Function Ay_XWh_Dist(A, Optional IgnCas As Boolean)
Ay_XWh_Dist = AyInto(AyCntDic(A, IgnCas).Keys, Ay_XCln(A))
End Function

Function Ay_XWh_DistFmt(A) As String()
Dim D As Dictionary
Set D = AyCntDic(A)
Ay_XWh_DistFmt = Dic_Fmt(D)
End Function

Function Ay_XWh_DistSy(A) As String()
Ay_XWh_DistSy = CvSy(Ay_XWh_Dist(A))
End Function

Function Ay_XWh_DistT1(A) As String()
Ay_XWh_DistT1 = Ay_XWh_Dist(Ay_XTak_T1(A))
End Function

Function Ay_XWh_Dup(A, Optional IgnCas As Boolean)
Dim D As Dictionary, I
Ay_XWh_Dup = Ay_XCln(A)
Set D = AyCntDic(A, IgnCas)
For Each I In AyNz(A)
    If D(I) > 1 Then
        Push Ay_XWh_Dup, I
    End If
Next
End Function

Function Ay_XWh_Eq3T(A, T1$, T2$, T3$) As String()
Ay_XWh_Eq3T = Ay_XWh_PredXABC(A, "LinHas3T", T1, T2, T3)
End Function

Function Ay_XWh_Fm(A, FmIx)
Dim O: O = A: Erase O
If 0 <= FmIx And FmIx <= UB(A) Then
    Dim J&
    For J = FmIx To UB(A)
        Push O, A(J)
    Next
End If
Ay_XWh_Fm = O
End Function

Function Ay_XWh_FmTo(A, FmIx, ToIx)
If Sz(A) = 0 Then Exit Function
Ay_XWh_FmTo = Ay_XCln(A)
FmToIxU_XAss FmIx, ToIx, UB(A)
Dim J&
For J = FmIx To ToIx
    Push Ay_XWh_FmTo, A(J)
Next
End Function

Function Ay_XWh_FstNEle(Ay, N&)
Dim O: O = Ay
ReDim Preserve O(N - 1)
Ay_XWh_FstNEle = O
End Function

Function Ay_XWh_FTIx(A, B As FTIx)
Ay_XWh_FTIx = Ay_XWh_FmTo(A, B.FmIx, B.ToIx)
End Function

Function Ay_XWh_XHas_Pfx(A, Pfx$) As String()
Ay_XWh_XHas_Pfx = Ay_XWh_Pfx(A, Pfx)
End Function

Function Ay_XWh_IxAy(A, IxAy)
Dim U&
    U = UB(A)
Ay_XWh_IxAy = Ay_XCln(A)
Dim Ix, J&
For Each Ix In AyNz(A)
    If 0 > Ix Or Ix > U Then
        XThw CSub, _
            "Given IxAy has some element not point to given Ay", _
            "[Ix from WhIxAy not point to Given-Ay] [   from ele ?] [Given Ay UB] [Given Ay] [Given WhIxAy]", _
            Ix, J, U, A, IxAy
    End If
    Push Ay_XWh_IxAy, A(Ix)
    J = J + 1
Next
End Function

Function Ay_XWh_Lik(A, Lik) As String()
Dim I
For Each I In AyNz(A)
    If I Like Lik Then PushI Ay_XWh_Lik, I
Next
End Function

Function Ay_XWh_LikAy(A, LikAy$()) As String()
Dim I, Lik
For Each I In AyNz(A)
    For Each Lik In LikAy
        If I Like Lik Then
            PushI Ay_XWh_LikAy, I
            Exit For
        End If
    Next
Next
End Function

Function Ay_XWh_Nm(A, B As WhNm) As String()
Ay_XWh_Nm = Ay_XExl_LikssAy(Ay_XWh_Re(A, B.Re), B.ExlAy)
End Function

Function Ay_XWh_NoEr(A, Msg$(), OEr$())
Dim J&
Erase OEr
If Not Ay_IsEqSz(A, Msg) Then Stop
For J = 0 To UB(A)
    If Msg(J) = "" Then Push Ay_XWh_NoEr, A(J) Else PushI OEr, Msg(J)
Next
End Function

Function Ay_XWh_NotPfx(A, Pfx$) As String()
Ay_XWh_NotPfx = Ay_XWh_PredXPNot(A, "XHas_Pfx", Pfx)
End Function

Function Ay_XWh_ObjPred(A, Obj, Pred$)
Dim I, O, X
Ay_XWh_ObjPred = Ay_XCln(A)
For Each I In AyNz(A)
    X = CallByName(Obj, Pred, VbMethod, I)
    If X Then
        Push Ay_XWh_ObjPred, I
    End If
Next
End Function

Function Ay_XWh_Patn(A, Patn$) As String()
If Sz(A) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then Ay_XWh_Patn = Ay_Sy(A): Exit Function
Ay_XWh_Patn = Ay_XWh_Re(A, New_Re(Patn))
End Function

Function Ay_XWh_PatnExl(A, Patn$, ExlLikss$) As String()
Ay_XWh_PatnExl = Ay_XExl_Likss(Ay_XWh_Patn(A, Patn), ExlLikss)
End Function
Function AyPatn_IxAy(A, Patn$) As Long()
AyPatn_IxAy = AyRe_IxAy(A, New_Re(Patn))
End Function
Function AyRe_IxAy(A, B As RegExp) As Long()
If Sz(A) = 0 Then Exit Function
Dim I, O&(), J&
For Each I In A
    If B.Test(I) Then PushI O, J
    J = J + 1
Next
AyRe_IxAy = O
End Function

Function Ay_XWh_Pfx(A, Pfx$) As String()
Dim I
For Each I In AyNz(A)
    If XHas_Pfx(I, Pfx) Then PushI Ay_XWh_Pfx, I
Next
End Function

Function Ay_XWh_Pred(A, Pred$)
Dim X
Ay_XWh_Pred = Ay_XCln(A)
For Each X In AyNz(A)
    If Run(Pred, X) Then
        Push Ay_XWh_Pred, X
    End If
Next
End Function

Function Ay_XWh_PredFalse(A, Pred$)
Dim X
Ay_XWh_PredFalse = Ay_XCln(A)
For Each X In AyNz(A)
    If Not Run(Pred, X) Then
        Push Ay_XWh_PredFalse, X
    End If
Next
End Function

Function Ay_XWh_PredNot(A, Pred$)
Ay_XWh_PredNot = Ay_XWh_PredFalse(A, Pred)
End Function

Function Ay_XWh_PredXAB(Ay, XAB$, A, B)
Dim X
Ay_XWh_PredXAB = Ay_XCln(Ay)
For Each X In AyNz(Ay)
    If Run(XAB, X, A, B) Then
        Push Ay_XWh_PredXAB, X
    End If
Next
End Function

Function Ay_XWh_PredXABC(Ay, XABC$, A, B, C)
Dim X
Ay_XWh_PredXABC = Ay_XCln(Ay)
For Each X In AyNz(Ay)
    If Run(XABC, X, A, B, C) Then
        Push Ay_XWh_PredXABC, X
    End If
Next
End Function

Function Ay_XWh_PredXAP(A, PredXAP$, ParamArray Ap())
Ay_XWh_PredXAP = Ay_XCln(A)
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In AyNz(A)
    Asg I, Av(0)
    If RunAv(PredXAP, Av) Then
        Push Ay_XWh_PredXAP, I
    End If
Next
End Function

Function Ay_XWh_PredXP(A, XP$, P)
Dim X
Ay_XWh_PredXP = Ay_XCln(A)
For Each X In AyNz(A)
    If Run(XP, X, P) Then
        Push Ay_XWh_PredXP, X
    End If
Next
End Function

Function Ay_XWh_PredXPNot(A, XP$, P)
Dim X
Ay_XWh_PredXPNot = Ay_XCln(A)
For Each X In AyNz(A)
    If Not Run(XP, X, P) Then
        Push Ay_XWh_PredXPNot, X
    End If
Next
End Function

Function Ay_XWh_Re(A, Re As RegExp) As String()
If IsNothing(Re) Then Ay_XWh_Re = Ay_Sy(A): Exit Function
Dim X
For Each X In AyNz(A)
    If Re.Test(X) Then PushI Ay_XWh_Re, X
Next
End Function
Function Ay_XWh_RmvEle(A, Ele)
Ay_XWh_RmvEle = Ay_XCln(A)
Dim I
For Each I In AyNz(A)
    If I <> Ele Then PushI Ay_XWh_RmvEle, I
Next
End Function
Function Ay_XWh_Rmv3T(A, T1$, T2$, T3$) As String()
Ay_XWh_Rmv3T = Ay_XRmv_3T(Ay_XWh_Eq3T(A, T1, T2, T3))
End Function

Function Ay_XWh_RmvT1(A, T1$) As String()
Ay_XWh_RmvT1 = Ay_XRmv_T1(Ay_XWh_T1(A, T1))
End Function

Function Ay_XWh_RmvTT(A, T1$, T2$) As String()
Ay_XWh_RmvTT = Ay_XRmv_TT(Ay_XWh_TT(A, T1, T2))
End Function

Function Ay_XWh_Sfx(A, Sfx$) As String()
Dim I
For Each I In AyNz(A)
    If XHas_Sfx(I, Sfx) Then PushI Ay_XWh_Sfx, I
Next
End Function

Function Ay_XWh_SingleEle(A)
Dim O: O = A: Erase O
Dim CntDry(): CntDry = AyCntDry(A)
If Sz(CntDry) = 0 Then
    Ay_XWh_SingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDry
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
Ay_XWh_SingleEle = O
End Function

Function Ay_XWh_Sng(A)
Ay_XWh_Sng = AyMinus(A, Ay_XWh_Dup(A))
End Function

Function Ay_XWh_SngEle(A)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = Ay_XCln(A)
Dim K, D As Dictionary
Set D = AyCntDic(A)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function Ay_XWh_T1(A, V) As String()
Ay_XWh_T1 = Ay_XWh_PredXP(A, "LinHasT1", V)
End Function

Function Ay_XWh_T1InAy(A, Ay$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), L
For Each L In A
    If Ay_XHas(Ay, Lin_T1(L)) Then Push O, L
Next
Ay_XWh_T1InAy = O
End Function

Function Ay_XWh_T1SelRst(A, T1) As String()
Dim L
For Each L In AyNz(A)
    If XShf_Term(L) = T1 Then PushI Ay_XWh_T1SelRst, L
Next
End Function

Function Ay_XWh_T2EqV(A$(), V) As String()
Ay_XWh_T2EqV = Ay_XWh_PredXP(A, "LinHasT2", V)
End Function

Function Ay_XWh_TT(A, T1$, T2$) As String()
Ay_XWh_TT = Ay_XWh_PredXAB(A, "LinHasTT", T1, T2)
End Function

Function Ay_XWh_TTSelRst(A, T1, T2) As String()
Dim L, X1$, X2$, Rst$
For Each L In AyNz(A)
    Lin_2TRstAsg L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI Ay_XWh_TTSelRst, Rst
        End If
    End If
Next
End Function

Function SyWhFmTo(A$(), FmIx, ToIx) As String()
Dim J&
For J = FmIx To ToIx
    Push SyWhFmTo, A(J)
Next
End Function

Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C As Boolean
Dim D&
Dim E As FTIx
Dim F$()
Dim G As WhNm
Dim H()
Dim I As RegExp
AyDupAss A, B
Ay_XWh_Dist A, C
Ay_XWh_DistFmt A
Ay_XWh_DistSy A
Ay_XWh_DistT1 A
Ay_XWh_Dup A
Ay_XWh_Eq3T A, B, B, B
Ay_XWh_Fm A, A
Ay_XWh_FmTo A, A, A
Ay_XWh_FstNEle A, D
Ay_XWh_FTIx A, E
Ay_XWh_XHas_Pfx A, B
Ay_XWh_IxAy A, A
Ay_XWh_Lik A, A
Ay_XWh_LikAy A, F
Ay_XWh_Nm A, G
Ay_XWh_NoEr A, F, F
Ay_XWh_NotPfx A, B
Ay_XWh_ObjPred A, A, B
Ay_XWh_Patn A, B
Ay_XWh_PatnExl A, B, B
Ay_XWh_Pfx A, B
Ay_XWh_Pred A, B
Ay_XWh_PredFalse A, B
Ay_XWh_PredNot A, B
Ay_XWh_PredXAB A, B, A, A
Ay_XWh_PredXABC A, B, A, A, A
Ay_XWh_PredXAP A, B, H
Ay_XWh_PredXP A, B, A
Ay_XWh_PredXPNot A, B, A
Ay_XWh_Re A, I
Ay_XWh_Rmv3T A, B, B, B
Ay_XWh_RmvT1 A, B
Ay_XWh_RmvTT A, B, B
Ay_XWh_Sfx A, B
Ay_XWh_SingleEle A
Ay_XWh_Sng A
Ay_XWh_SngEle A
Ay_XWh_T1 A, A
Ay_XWh_T1InAy A, F
Ay_XWh_T1SelRst A, A
Ay_XWh_T2EqV F, A
Ay_XWh_TT A, B, B
Ay_XWh_TTSelRst A, A, A
SyWhFmTo F, A, A
End Sub

Private Sub Z()
End Sub
