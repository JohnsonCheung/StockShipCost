Attribute VB_Name = "MVb_Ay__Oy"
Option Compare Binary
Option Explicit

Sub Oy_XDo(Oy, DoFun$)
Dim O
For Each O In Oy
    Run DoFun, O ' DoFunNm cannot be like a Excel.Address (eg, A1, XX1)
Next
End Sub

Sub Oy_XDoMth(A, Mth$)
Dim J&
For J = 0 To UB(A)
    CallByName A(J), Mth, VbMethod
Next
End Sub

Sub Oy_XDoSub_P(A, SubNm$, P)
If Sz(A) = 0 Then Exit Sub
Dim O
For Each O In AyNz(A)
    CallByName O, SubNm, VbMethod, P
Next
End Sub

Function Oy_Fst_PrpEqV(A, P, V)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Obj_Prp(X, P) = V Then Asg X, Oy_Fst_PrpEqV: Exit Function
Next
End Function

Function OyHas(A, Obj) As Boolean
Dim X, Op&
Op = ObjPtr(Obj)
For Each X In AyNz(A)
    If ObjPtr(X) = Op Then OyHas = True: Exit Function
Next
End Function

Function OyMap(A, MapMthNm$) As Variant()
OyMap = OyMapInto(A, MapMthNm, EmpAy)
End Function

Function OyMapInto(A, MapFunNm$, OIntoAy)
Dim Obj, J&, U&
U = UB(A)
Dim O
O = OIntoAy
ReSz O, U
For J = 0 To U
    Asg Run(MapFunNm, A(J)), O(J)
Next
OyMapInto = O
End Function

Function OyPrpAy(Oy, PrpNm) As Variant()
OyPrpAy = OyPrpAyInto(Oy, PrpNm, EmpAy)
End Function

Function OyPrpAyInto(Oy, PrpNm, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(Oy) > 0 Then
    Dim I
    For Each I In Oy
        Push O, Obj_Prp(I, PrpNm)
    Next
End If
OyPrpAyInto = O
End Function

Function OyPrpIntAy(A, PrpNm$) As Integer()
OyPrpIntAy = OyPrpInto(A, PrpNm, EmpIntAy)
End Function

Function OyPrpInto(A, PrpNm$, OInto)
If Sz(A) = 0 Then OyPrpInto = Ay_XCln(OInto): Exit Function
OyPrpInto = ItrPrp_Into(A, PrpNm, OInto)
End Function

Function OyPrpSrtedUniqAy(A, PrpNm$) As Variant()
OyPrpSrtedUniqAy = Ay_XSrt(Ay_XWh_Dist(OyPrpAy(A, PrpNm)))
End Function

Function OyPrpSrtedUniqIntAy(A, PrpNm$) As Integer()
OyPrpSrtedUniqIntAy = Ay_XSrt(Ay_XWh_Dist(OyPrpIntAy(A, PrpNm)))
End Function

Function OyPrpSrtedUniqSy(A, PrpNm$) As Variant()
OyPrpSrtedUniqSy = Ay_XSrt(Ay_XWh_Dist(Oy_PrpSy(A, PrpNm)))
End Function

Function Oy_PrpSy(A, PrpNm$) As String()
Oy_PrpSy = OyPrpInto(A, PrpNm, EmpSy)
End Function

Function OyRmvFstNEle(A, N&)
Dim O: O = A
ReDim O(N - 1)
Dim J&
For J = 0 To UB(A) - N
    Set O(J) = A(N + J)
Next
OyRmvFstNEle = O
End Function

Function Oy_XWh_NoNothing(A)
Oy_XWh_NoNothing = Ay_XCln(A)
Dim I
For Each I In AyNz(A)
    If Not IsNothing(I) Then PushObj Oy_XWh_NoNothing, I
Next
End Function

Function OySrt_By_CompoundPrp(A, PrpSsl$)
Dim O: O = A: Erase O
Dim Sy$(): Sy = Oy_PrpSy(A, PrpSsl)
Dim Ix&(): Ix = Ay_XSrt_IntoIxAy(Sy)
Dim J&
For J = 0 To UB(Ix)
    PushObj O, A(Ix(J))
Next
OySrt_By_CompoundPrp = O
End Function

Function OyToStr$(A)
Dim O$(), I
For Each I In A
    Push O, CallByName(I, "ToStr", VbGet)
Next
OyToStr = JnCrLf(O)
End Function

Function OyWhIxAy(A, IxAy)
Dim O: O = A: Erase O
Dim U&: U = UB(IxAy)
Dim J&
ReSz O, U
For J = 0 To U
    Asg A(IxAy(J)), O(J)
Next
OyWhIxAy = O
End Function

Function OyWhIxSelIntPrp(A, WhIx, PrpNm$) As Integer()
OyWhIxSelIntPrp = OyWhIxSelPrp(A, WhIx, PrpNm, EmpIntAy)
End Function

Function OyWhIxSelPrp(A, WhIx, PrpNm$, OupAy)
Dim Oy1: Oy1 = OyWhIxAy(A, WhIx)  ' Oy1 is subset of Oy
OyWhIxSelPrp = OyPrpInto(Oy1, PrpNm, OupAy)
End Function

Function OyWhIxSelSyPrp(A, WhIx, PrpNm$) As String()
OyWhIxSelSyPrp = OyWhIxSelPrp(A, WhIx, PrpNm, EmpSy)
End Function

Function OyWhNm(A, B As WhNm)
Dim X
For Each X In AyNz(A)
    If Nm_IsSel(X.Name, B) Then PushObj OyWhNm, X
Next
End Function

Function OyWhNmExl(A, ExlAy$)
If ExlAy = "" Then OyWhNmExl = A: Exit Function
Dim X, LikAy$(), O
O = A
Erase O
LikAy = Ssl_Sy(ExlAy)
For Each X In AyNz(A)
    If Not IsInLikAy(X.Name, LikAy) Then PushObj O, X
Next
OyWhNmExl = O
End Function

Function Oy_XWh_NmHasPfx(A, Pfx$)
Oy_XWh_NmHasPfx = Oy_XWh_PredXP(A, "Any_Obj_NmPfx", Pfx)
End Function

Function OyWhNmPatn(A, Patn$)
If Patn = "." Then OyWhNmPatn = A: Exit Function
Dim X, O, Re As New RegExp
O = A
Erase O
Re.Pattern = Patn
For Each X In AyNz(A)
    If Re.Test(X.Name) Then PushObj O, X
Next
OyWhNmPatn = O
End Function

Function OyWhNmPatnExl(A, Patn$, ExlAy$)
OyWhNmPatnExl = OyWhNmExl(OyWhNmPatn(A, Patn), ExlAy)
End Function

Function OyWhNmReExl(A, Re As RegExp, ExlLikAy$())
If Sz(A) = 0 Then OyWhNmReExl = A: Exit Function
Dim X
For Each X In AyNz(A)
    If Nm_IsSel_ByReExl(X.Name, Re, ExlLikAy) Then PushObj OyWhNmReExl, X
Next
End Function

Function Oy_XWh_PredXP(A, XP$, P)
Dim O, X
O = A
Erase O
For Each X In AyNz(A)
    If Run(XP, X, P) Then
        PushObj A, X
    End If
Next
Oy_XWh_PredXP = O
End Function

Function OyWhPrp(A, PrpNm$, PrpEqToVal)
Dim O
   O = A
   Erase O
If Not Sz(A) > 0 Then
   Dim I
   For Each I In A
       If CallByName(I, PrpNm, VbGet) = PrpEqToVal Then PushObj O, I
   Next
End If
End Function

Function OyWhPrpEqV(A, P, V)
Dim X
If Sz(A) > 0 Then
    For Each X In AyNz(A)
        If Obj_Prp(X, P) = V Then
            Set OyWhPrpEqV = X: Exit Function
        End If
    Next
End If
Set OyWhPrpEqV = Nothing
End Function

Function OyWhPrpEqValSelPrpInt(A, WhPrpNm$, EqVal, SelPrpNm$) As Integer()
Dim Oy1: Oy1 = OyWhPrpEqV(A, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpInt = OyPrpIntAy(Oy1, SelPrpNm)
End Function

Function OyWhPrpEqValSelPrpSy(A, WhPrpNm$, EqVal, SelPrpNm$) As String()
Dim Oy1: Oy1 = OyWhPrpEqV(A, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpSy = Oy_PrpSy(Oy1, SelPrpNm)
End Function

Function OyWhPrpIn(A, P, InAy)
Dim X, O
If Sz(A) = 0 Or Sz(InAy) Then OyWhPrpIn = A: Exit Function
O = A
Erase O
For Each X In AyNz(A)
    If Ay_XHas(InAy, Obj_Prp(X, P)) Then PushObj O, X
Next
OyWhPrpIn = O
End Function

Function SelOy(A, PrpSsl$) As Variant()

End Function

Private Sub ZZ_OyDrs()
'Ws_XVis DrsNew_Ws(OyDrs(CurrentDb.TableDefs("ZZ_Dbt_XUpd_Seq").Fields, "Name Type OrdinalPosition"))
End Sub

Private Sub ZZ_OyPrpAy()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CurPj_X.MdAy).PrpAy("CodePane", CdPanAy)
Stop
End Sub
