Attribute VB_Name = "MVb_Ay_Fst"
Option Compare Binary
Option Explicit

Function Ay_FstEle(Ay)
If Sz(Ay) = 0 Then Exit Function
Asg Ay(0), Ay_FstEle
End Function

Function Ay_FstEqV(A, V)
If Ay_XHas(A, V) Then Ay_FstEqV = V
End Function

Function Ay_FstLik$(A, Lik$)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In A
    If X Like Lik Then Ay_FstLik = X: Exit Function
Next
End Function

Function Ay_FstLikItm(A, Lik, Itm)
Ay_FstLikItm = Ay_FstPredXABYes(A, "LinHasLikItm", Lik, Itm)
End Function

Function Ay_FstNEle(A, N)
Dim O: O = A
ReDim Preserve O(N - 1)
Ay_FstNEle = O
End Function

Function Ay_FstPfx$(PfxAy, Lin$)
Dim P
For Each P In PfxAy
    If XHas_Pfx(Lin, CStr(P)) Then Ay_FstPfx = P: Exit Function
Next
End Function

Function Ay_FstPredPX(A, PX$, P)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Run(PX, P, X) Then Asg X, Ay_FstPredPX: Exit Function
Next
End Function

Function Ay_FstPredXABYes(Ay, XAB$, A, B)
Dim X
For Each X In AyNz(Ay)
    If Run(XAB, X, A, B) Then Asg X, Ay_FstPredXABYes: Exit Function
Next
End Function

Function Ay_FstPredXP(A, XP$, P)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Run(XP, X, P) Then Asg X, Ay_FstPredXP: Exit Function
Next
End Function

Function Ay_FstT1$(A, T1, Optional CasSen As Boolean)
Dim L
For Each L In AyNz(A)
    If Lin_T1(L) = T1 Then Ay_FstT1 = L: Exit Function
Next
End Function

Function Ay_FstRmvT1$(A, T1)
Ay_FstRmvT1 = RmvT1(Ay_FstT1(A, T1))
End Function

Function Ay_FstT2$(A, T2)
Ay_FstT2 = Ay_FstPredXP(A, "LinHasT2", T2)
End Function

Function Ay_FstTT$(A, T1, T2)
Ay_FstTT = Ay_FstPredXABYes(A, "LinHasTT", T1, T2)
End Function


Function Ay_FstRmvTT$(A, T1$, T2$)
Dim X, X1$, X2$, Rst$
For Each X In AyNz(A)
    Lin_2TRstAsg X, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            Ay_FstRmvTT = X
            Exit Function
        End If
    End If
Next
End Function


