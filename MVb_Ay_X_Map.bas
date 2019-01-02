Attribute VB_Name = "MVb_Ay_X_Map"
Option Compare Binary
Option Explicit

Function AyMap(A, Map$)
AyMap = AyMap_Into(A, Map, EmpAy)
End Function

Function AyMapABX_Into(Ay, ABX$, A, B, OInto)
Dim O: O = OInto
Erase O
If Sz(Ay) > 0 Then
    Dim J&, X
    ReDim O(UB(A))
    For Each X In AyNz(A)
        Asg Run(ABX, A, B, X), O(J)
        J = J + 1
    Next
End If
AyMapABX_Into = O
End Function

Function AyMapABX_Sy(Ay, ABX$, A, B) As String()
AyMapABX_Sy = AyMapABX_Into(Ay, ABX, A, B, EmpSy)
End Function

Function AyMapAXB_Into(Ay, AXB$, A, B, OInto)
Dim O: O = OInto
Erase O
If Sz(Ay) > 0 Then
    Dim J&, X
    ReDim O(UB(A))
    For Each X In AyNz(A)
        Asg Run(AXB, A, X, B), O(J)
        J = J + 1
    Next
End If
AyMapAXB_Into = O
End Function

Function AyMapAXB_Sy(Ay, AXB$, A, B)
AyMapAXB_Sy = AyMapAXB_Into(Ay, AXB, A, B, EmpSy)
End Function

Function AyMapAsgAy(A, OAy, MthNm$, ParamArray Ap())
If Sz(A) = 0 Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O
O = OAy
Erase O
Dim U&: U = UB(A)
    ReDim O(U)
For Each I In A
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMapAsgAy = O
End Function

Function AyMapAsgSy(A, MthNm$, ParamArray Ap()) As String()
If Sz(A) = 0 Then Exit Function
Dim Av(): Av = Ap
If Sz(Av) = 0 Then
    AyMapAsgSy = AyMap_Sy(A, MthNm)
    Exit Function
End If
Dim I, J&
Dim O$()
    ReDim O(UB(A))
    Av = AyIns(Av)
    For Each I In A
        Asg I, Av(0)
        Asg RunAv(MthNm, Av), O(J)
        J = J + 1
    Next
AyMapAsgSy = O
End Function


Function AyMap_Into(A, MapFunNm$, OInto)
Dim O: O = Ay_XCln(OInto)
Dim I
If Sz(A) > 0 Then
    For Each I In A
        Push O, Run(MapFunNm, I)
    Next
End If
AyMap_Into = O
End Function

Function AyMap_LngAy(A, MapMthNm$) As Long()
AyMap_LngAy = AyMap_Into(A, MapMthNm, EmpLngAy)
End Function

Function AyMapXP_Ay(A, XP$, P) As Variant()
AyMapXP_Ay = AyMapPX_Into(A, XP, P, EmpAy)
End Function

Function AyMapPX_Into(A, PX$, P, OInto)
AyMapPX_Into = Ay_XCln(OInto)
Dim X
For Each X In AyNz(A)
    Push AyMapPX_Into, Run(PX, P, X)
Next
End Function

Function AyMapPX_Sy(A, PX$, P) As String()
AyMapPX_Sy = AyMapPX_Into(A, PX, P, EmpSy)
End Function

Function AyMap_Sy(A, MapMthNm$) As String()
AyMap_Sy = AyMap_Into(A, MapMthNm, EmpSy)
End Function

Function AyMapXABCD_Into(Ay, XABCD$, A, B, C, D, OInto)
Erase OInto
If Sz(Ay) = 0 Then AyMapXABCD_Into = OInto: Exit Function
Dim X
For Each X In AyNz(A)
    Push OInto, Run(XABCD, X, A, B, C, D)
Next
AyMapXABCD_Into = OInto
End Function

Function AyMapXABC_Into(Ay, XABC$, A, B, C, OInto)
Erase OInto
Dim X
For Each X In AyNz(A)
    Push OInto, Run(XABC, X, A, B, C)
Next
AyMapXABC_Into = OInto
End Function

Function AyMapXAB_Into(Ay, XAB$, A, B, OInto)
AyMapXAB_Into = Ay_XCln(OInto)
Dim X
For Each X In AyNz(Ay)
    Push AyMapXAB_Into, Run(XAB, X, A, B)
Next
End Function

Function AyMapXAB_Sy(Ay, XAB$, A, B) As String()
AyMapXAB_Sy = AyMapXAB_Into(Ay, XAB, A, B, EmpSy)
End Function

Function AyMapXP_Into(A, XP$, P, OInto)
Dim O, X
O = OInto
Erase O
For Each X In AyNz(A)
    Push O, Run(XP, X, P)
Next
AyMapXP_Into = O
End Function

Function AyMapXP_Sy(A, XP$, P) As String()
AyMapXP_Sy = AyMapXP_Into(A, XP, P, EmpSy)
End Function

Private Sub ZZ_AyMap()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Ass Sz(Act) = 4
Ass Act(0) = 2
Ass Act(1) = 4
Ass Act(2) = 6
Ass Act(3) = 8
End Sub

Private Sub ZZ_AyMap_Sy()
Dim Ay$(): Ay = AyMap_Sy(Array("skldfjdf", "aa"), "XRmv_FstChr")
Stop
End Sub

Private Sub Z_AyMap_Sy()
Dim Ay$(): Ay = AyMap_Sy(Array("skldfjdf", "aa"), "XRmv_FstChr")
Stop
End Sub

Private Sub Z_AyMapXAP()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Ass Sz(Act) = 4
Ass Act(0) = 2
Ass Act(1) = 4
Ass Act(2) = 6
Ass Act(3) = 8
End Sub


Private Sub Z()
Z_AyMap_Sy
Z_AyMapXAP
MVb_Ay_X_Map:
End Sub
