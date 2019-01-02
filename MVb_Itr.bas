Attribute VB_Name = "MVb_Itr"
Option Compare Binary
Option Explicit

Function Itr_Ay(A)
Itr_Ay = Itr_Into(A, Array())
End Function

Function Itr_XAdd_Sfx(A, Sfx$) As String()
Dim X
For Each X In A
    Push Itr_XAdd_Sfx, X & Sfx
Next
End Function

Function Itr_XAdd_Pfx(A, Pfx$) As String()
Dim X
For Each X In A
    Push Itr_XAdd_Pfx, Pfx & X
Next
End Function

Function Itr_ClnAy(A)
If A.Count = 0 Then Exit Function
Dim X
For Each X In A
    Itr_ClnAy = Array(X)
    Exit Function
Next
End Function

Function Itr_Cnt_PrpTrue(A, BoolPrpNm)
Dim O&, X
For Each X In A
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
Itr_Cnt_PrpTrue = O
End Function

Sub Itr_XDo(A, DoFun$)
Dim I
For Each I In A
    Run DoFun, I
Next
End Sub

Sub Itr_XDo_ObjMthNm(A, ObjMthNm$)
Dim I
For Each I In A
    CallByName I, ObjMthNm, VbMethod
Next
End Sub

Sub Itr_XDo_PX(A, PX$, P)
Dim X
For Each X In A
    Run PX, P, X
Next
End Sub

Sub Itr_XDo_XP(A, XP$, P)
Dim X
For Each X In A
    Run XP, X, P
Next
End Sub

Function Itr_FstItm(A)
Dim X
For Each X In A
    Asg X, Itr_FstItm
    Exit Function
Next
XThw CSub, "No itm in Itr", "Itr-TypeName", TypeName(A)
End Function

Function Itr_FstItmNm(A, Nm) ' Return first element in Itr-A with name eq Nm
Dim X
For Each X In A
    If Obj_Nm(X) = Nm Then Set Itr_FstItmNm = X: Exit Function
Next
Set Itr_FstItmNm = Nothing
End Function

Function Itr_FstItm_PredXP(A, XP$, P)
Dim X
For Each X In A
    If Run(XP, X, P) Then Asg X, Itr_FstItm_PredXP: Exit Function
Next
End Function

Function Itr_FstItm_PrpEqV(A, P, V) 'Return first element in Itr-A with its Prp-P eq to V
Dim X
For Each X In A
    If Obj_Prp(X, P) = V Then Set Itr_FstItm_PrpEqV = X: Exit Function
Next
Set Itr_FstItm_PrpEqV = Nothing
End Function

Function Itr_FstItm_PrpTrue(A, P) 'Return first element in Itr-A with its Prp-P being true
Dim X
For Each X In A
    If Obj_Prp(X, P) Then Set Itr_FstItm_PrpTrue = X: Exit Function
Next
Set Itr_FstItm_PrpTrue = Nothing
End Function

Function Itr_XHas_Itm(A, Itm) As Boolean
Dim I
For Each I In A
    If I = Itm Then Itr_XHas_Itm = True: Exit Function
Next
End Function

Function Itr_XHas_Nm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then Itr_XHas_Nm = True: Exit Function
Next
End Function

Function Itr_XHas_Nm_WhRe(A, Re As RegExp) As Boolean
Dim I
For Each I In A
    If Re.Test(I.Name) Then Itr_XHas_Nm_WhRe = True: Exit Function
Next
End Function

Function Itr_XHas_PrpVal(A, P, V) As Boolean
Dim X
For Each X In A
    If Obj_Prp(X, P) = V Then Itr_XHas_PrpVal = True: Exit Function
Next
End Function

Function Itr_XHas_PrpTrue(A, P) As Boolean
Dim X
For Each X In A
    If Obj_Prp(X, P) Then Itr_XHas_PrpTrue = True: Exit Function
Next
End Function


Function Itr_Into(A, OInto)
Itr_Into = Ay_XCln(OInto)
Dim X
For Each X In A
    Push Itr_Into, X
Next
End Function

Function ItrIsEqNm(A, B)
ItrIsEqNm = Ay_IsSam(Itr_Ny(A), Itr_Ny(B))
End Function

Function ItrMap(A, Map$) As Variant()
ItrMap = Itr_Into_ByMap(A, Map, EmpAy)
End Function

Function Itr_Into_ByMap(A, Map$, OInto)
Dim O: O = OInto
Erase O
Dim X
For Each X In A
    Push O, Run(Map, X)
Next
Itr_Into_ByMap = O
End Function

Function Itr_Sy_ByMap(A, Map$) As String()
Itr_Sy_ByMap = Itr_Into_ByMap(A, Map, EmpSy)
End Function

Function Itr_MaxPrp(A, P)
Dim X, O
For Each X In A
    O = Max(O, Obj_Prp(X, P))
Next
Itr_MaxPrp = O
End Function
Function Oy_Ny(A) As String()
Dim I
For Each I In AyNz(A)
    PushI Oy_Ny, Obj_Nm(I)
Next
End Function
Function Itr_Ny(A) As String()
Dim I
For Each I In A
    PushI Itr_Ny, Obj_Nm(I)
Next
End Function

Function Itr_Ny_WhNmLik(A, Lik) As String()
Itr_Ny_WhNmLik = Ay_XWh_Lik(Itr_Ny(A), Lik)
End Function

Function Itr_Ny_WH_PATN_EXL(A, Optional Patn$, Optional Exl$) As String()
Itr_Ny_WH_PATN_EXL = Ay_XWh_PatnExl(Itr_Ny(A), Patn, Exl)
End Function

Function ItrPred_IsAllFalse(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then Exit Function
Next
ItrPred_IsAllFalse = True
End Function

Function ItrPred_IsAllTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then Exit Function
Next
ItrPred_IsAllTrue = True
End Function

Function ItrPred_IsSomFalse(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then ItrPred_IsSomFalse = True: Exit Function
Next
End Function

Function ItrPred_IsSomTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then ItrPred_IsSomTrue = True: Exit Function
Next
End Function

Function Itr_XHas_PredTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then Itr_XHas_PredTrue = True: Exit Function
Next
End Function

Function ItrPrp_Ay(A, P) As Variant()
ItrPrp_Ay = ItrPrp_Into(A, P, EmpAy)
End Function

Function ItrPrp_Into(A, P, OInto)
ItrPrp_Into = Ay_XCln(OInto)
Dim I
For Each I In A
    Push ItrPrp_Into, Obj_Prp(I, P)
Next
End Function

Function ItrPrp_Sy(A, P) As String()
ItrPrp_Sy = ItrPrp_Into(A, P, EmpSy)
End Function

Function Itr_Vy(A) As Variant()
Itr_Vy = ItrPrp_Ay(A, "Value")
End Function

Function Itr_XWh_Nm(A, B As WhNm)
Itr_XWh_Nm = Itr_XWh_NmInto(A, B, EmpAy)
End Function

Function Itr_XWh_NmInto(A, B As WhNm, OInto)
Itr_XWh_NmInto = Ay_XCln(OInto)
Dim X
For Each X In A
    If Nm_IsSel(X.Name, B) Then PushObj Itr_XWh_NmInto, X
Next
End Function

Function Itr_XWh_NmReExl(A, Re As RegExp, ExlLikAy$())
Itr_XWh_NmReExl = Itr_XWh_NmReExlInto(A, Re, ExlLikAy, EmpAy)
End Function

Function Itr_XWh_NmReExlInto(A, Re As RegExp, ExlAy$(), OInto)
Dim X
Itr_XWh_NmReExlInto = Ay_XCln(OInto)
For Each X In A
    If Nm_IsSel_ByReExl(X.Name, Re, ExlAy) Then PushObj Itr_XWh_NmReExlInto, X
Next
End Function

Function Itr_Into_WhInNy(A, InNy$(), OInto)
Dim O, X
O = OInto
Erase O
For Each X In A
    If Ay_XHas(InNy, X.Name) Then PushObj O, X
Next
Itr_Into_WhInNy = O
End Function

Function Itr_XWh_PredPrpTrue_Ay(A, Pred$, P)
Itr_XWh_PredPrpTrue_Ay = Itr_PrpInto_WhItmPrepTrue(A, Pred, P, EmpAy)
End Function

Function Itr_PrpInto_WhItmPrepTrue(A, Pred$, P, OInto)
Dim O: O = OInto
Erase O
Dim X
For Each X In A
    If Run(Pred, X) Then
        Push O, Obj_Prp(X, P)
    End If
Next
Itr_PrpInto_WhItmPrepTrue = O
End Function

Function Itr_PrpSy_WhPrepPrpTrue(A, Pred$, P) As String()
Itr_PrpSy_WhPrepPrpTrue = Itr_PrpInto_WhItmPrepTrue(A, Pred, P, EmpSy)
End Function

Function Itr_XWh_PrpEqV(A, P, V)
Dim O: O = Itr_ClnAy(A): If IsEmpty(O) Then Exit Function
Dim X
For Each X In A
    If Obj_Prp(X, P) = V Then PushObj O, X
Next
Itr_XWh_PrpEqV = O
End Function

Function Itr_XWh_PrpTrue(A, P)
Itr_XWh_PrpTrue = Itr_XWh_PrpTrue_Into(A, P, EmpAy)
End Function

Function Itr_XWh_PrpTrue_Into(A, P, OInto)
Dim O: O = OInto: Erase O
Dim X
For Each X In A
    If Obj_Prp(A, P) Then
        Push O, X
    End If
Next
Itr_XWh_PrpTrue_Into = O
End Function

Function Itr_Ay_WhNm(A, B As WhNm)
Itr_Ay_WhNm = Itr_XWh_NmInto(A, B, EmpAy)
End Function

Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C As RegExp
Dim D$()
Dim E As WhNm
Itr_Ay A
Itr_ClnAy A
Itr_XDo A, B
Itr_XDo A, B
Itr_XDo_PX A, B, A
Itr_XDo_XP A, B, A
Itr_FstItm A
Itr_FstItm A
Itr_FstItmNm A, A
Itr_FstItm_PredXP A, B, A
Itr_FstItm_PrpEqV A, A, A
Itr_FstItm_PrpTrue A, A
Itr_XHas_Itm A, A
Itr_XHas_Nm A, A
Itr_XHas_Nm_WhRe A, C
Itr_XHas_PrpVal A, A, A
Itr_XHas_PrpTrue A, A
Itr_Into A, A
ItrIsEqNm A, A
ItrMap A, B
Itr_Into_ByMap A, B, A
Itr_Sy_ByMap A, B
Itr_MaxPrp A, A
Itr_Ny A
ItrPred_IsAllFalse A, B
ItrPred_IsAllTrue A, B
ItrPred_IsSomFalse A, B
Itr_XHas_PredTrue A, B
ItrPrp_Ay A, A
ItrPrp_Into A, A, A
ItrPrp_Sy A, A
Itr_Vy A
Itr_Into_WhInNy A, D, A
Itr_XWh_Nm A, E
Itr_XWh_NmInto A, E, A
Itr_XWh_NmReExl A, C, D
Itr_XWh_NmReExlInto A, C, D, A
Itr_Into_WhInNy A, D, A
Itr_XWh_PredPrpTrue_Ay A, B, A
Itr_PrpInto_WhItmPrepTrue A, B, A, A
Itr_PrpSy_WhPrepPrpTrue A, B, A
Itr_XWh_PrpEqV A, A, A
Itr_XWh_PrpTrue A, A
Itr_XWh_PrpTrue_Into A, A, A
Itr_Ay_WhNm A, E
End Sub

Private Sub Z()
End Sub
