Attribute VB_Name = "MVb_Ay_Shf"
Option Compare Binary
Option Explicit

Private Sub Z_Ay_XShf_()
Dim Ay(), Exp, Act, ExpAyAft()
Ay = Array(1, 2, 3, 4)
Exp = 1
ExpAyAft = Array(2, 3, 4)
GoSub Tst
Exit Sub
Tst:
Act = Ay_XShf_(Ay)
Debug.Assert IsEq(Exp, Act)
Debug.Assert Ay_IsEq(Ay, ExpAyAft)
Return
End Sub

Private Sub Z_Ay_XShf_Itm()
Dim OAy$(), Itm, EptOAy
Ept = "1"
Itm = "AA"
OAy = Ap_Sy("AA=1")
EptOAy = Ap_Sy("")
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XShf_Itm(OAy, Itm)
    C
    Ass Ay_IsEq(OAy, EptOAy)
    Return
End Sub

Private Sub Z_Ay_XShf_ItmNy()
Dim A$(), ItmNy0
A = Ssl_Sy("Req Dft=ABC VTxt=kkk")
ItmNy0 = "Req ABC VTxt"
Ept = Array("Req", "", "kkk", Ap_Sy("Dft=ABC"))
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XShf_ItmNy(A, ItmNy0)
    C
    Return
End Sub


Function Ay_XShf_(OAy)
Ay_XShf_ = OAy(0)
OAy = Ay_XExl_FstEle(OAy)
End Function

Function Ay_XShf_FstNEle(OAy, N)
Ay_XShf_FstNEle = Ay_FstNEle(OAy, N)
OAy = Ay_XExl_FstNEle(OAy, N)
End Function

Function Ay_XShf_Itm$(OAy, Itm)
Dim J%
If XTak_FstChr(Itm) = "?" Then Ay_XShf_Itm = Ay_XShf_QItm(OAy, XRmv_FstChr(Itm)): Exit Function
For J = 0 To UB(OAy)
    If XHas_Pfx(OAy(J), Itm) Then
        Ay_XShf_Itm = XBrk(OAy(J), "=")(1)
        OAy = Ay_XExl_EleAt(OAy, J)
        Exit Function
    End If
Next
End Function

Function Ay_XShf_ItmEq(A, Itm$) As Variant()
Dim B$
    Dim Lik$
    Lik = Itm & "=*"
    B = Ay_FstLik(A, Lik)
If B = "" Then
    Ay_XShf_ItmEq = Array("", A)
Else
    Ay_XShf_ItmEq = Array(Trim(XRmv_Pfx(B, Itm & "=")), Ay_XExl_EleLik(A, Lik))
End If
End Function

Function Ay_XShf_ItmNy(A$(), ItmNy0) As Variant()
Dim Ny$(), A1$()
    Ny = CvNy(ItmNy0)
    A1 = A
Dim O() As Variant, Ay(), J%
ReDim O(Sz(Ny))
For J = 0 To UB(Ny)
    Ay = Ay_XShf_ItmEq(A1, Ny(J))
    O(J) = Ay(0)
    A1 = Ay(1)
Next
O(Sz(Ny)) = Ay(1)
Ay_XShf_ItmNy = O
End Function

Function Ay_XShf_QItm$(OAy, QItm)
Dim I, J%
For Each I In AyNz(OAy)
    If QItm = I Then Ay_XShf_QItm = QItm: OAy = Ay_XExl_EleAt(OAy, J): Exit Function
    J = J + 1
Next
End Function

Function Ay_XShf_Star(OAy, OItmy$()) As String()
Dim NStar%: NStar = AyNPfxStar(OItmy)
Ay_XShf_Star = Ay_XShf_FstNEle(OAy, NStar)
OItmy = Ay_XExl_FstNEle(OItmy, NStar)
End Function


Private Sub Z()
Z_Ay_XShf_
Z_Ay_XShf_Itm
Z_Ay_XShf_ItmNy
MVb_Ay_XShf_:
End Sub
