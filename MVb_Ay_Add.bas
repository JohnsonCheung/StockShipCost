Attribute VB_Name = "MVb_Ay_Add"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Ay_Add."

Function Ay_XAdd_1(A)
Ay_XAdd_1 = Ay_XAdd_N(A, 1)
End Function

Function Ay_XAdd_(A, B)
Ay_XAdd_ = A
PushAy Ay_XAdd_, B
End Function

Function AyAp_XAdd(Ay, ParamArray Itm_or_Ay_Ap())
Const CSub$ = CMod & "AyAp_XAdd"
Dim Av(): Av = Itm_or_Ay_Ap
If Not IsArray(Ay) Then XThw CSub, "Fst parameter must be array", "Fst-Pm-TyeName", TypeName(Ay)
Dim I
AyAp_XAdd = Ay
For Each I In Av
    If IsArray(I) Then
        PushAy AyAp_XAdd, I
    Else
        Push AyAp_XAdd, I
    End If
Next
End Function

Function Ay_XAdd_FunCol(A, FunNm$) As Variant()
Dim X
For Each X In AyNz(A)
    PushI Ay_XAdd_FunCol, Array(X, Run(FunNm, X))
Next
End Function

Function Ay_XAdd_Itm(A, Itm)
Dim O
O = A
Push O, Itm
Ay_XAdd_Itm = O
End Function

Function Ay_XAdd_N(A, N)
Ay_XAdd_N = Ay_XCln(A)
Dim X
For Each X In AyNz(A)
    PushI Ay_XAdd_N, X + N
Next
End Function

Private Sub Z_Ay_XAdd_()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = Ay_XAdd_(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
Ay_XAss_Eq Exp, Act
Ay_XAss_Eq Ay1, Array(1, 2, 2, 2, 4, 5)
Ay_XAss_Eq Ay2, Array(2, 2)
End Sub


Private Sub ZZ_Ay_XAdd_()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = Ay_XAdd_(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
Ay_IsEq_XAss Exp, Act
Ay_IsEq_XAss Ay1, Array(1, 2, 2, 2, 4, 5)
Ay_IsEq_XAss Ay2, Array(2, 2)
End Sub

Private Sub ZZ_Ay_XAdd_Pfx()
Dim A, Act$(), Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Exp = Ap_Sy("* 1", "* 2", "* 3", "* 4")
GoSub Tst
Exit Sub
Tst:
Act = Ay_XAdd_Pfx(A, Pfx)
Debug.Assert Ay_IsEq(Act, Exp)
Return
End Sub

Private Sub ZZ_Ay_XAdd_PfxSfx()
Dim A, Act$(), Sfx$, Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = Ap_Sy("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = Ay_XAdd_PfxSfx(A, Pfx, Sfx)
Debug.Assert Ay_IsEq(Act, Exp)
Return
End Sub

Function AyTab(A) As String()
AyTab = Ay_XAdd_Pfx(A, vbTab)
End Function

Private Sub ZZ_Ay_XAdd_Sfx()
Dim A, Act$(), Sfx$, Exp$()
A = Array(1, 2, 3, 4)
Sfx = "#"
Exp = Ap_Sy("1#", "2#", "3#", "4#")
GoSub Tst
Exit Sub
Tst:
Act = Ay_XAdd_Sfx(A, Sfx)
Debug.Assert Ay_IsEq(Act, Exp)
Return
End Sub


Private Sub Z()
Z_Ay_XAdd_
MVb_Ay_Add:
End Sub
