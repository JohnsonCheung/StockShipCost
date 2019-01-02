Attribute VB_Name = "MVb_Ay_Has"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Ay_XHas."
Function Ay_XHas(A, M) As Boolean
Dim I
For Each I In AyNz(A)
    If I = M Then Ay_XHas = True: Exit Function
Next
End Function

Function Ay_XHasAy(A, Ay) As Boolean
Dim I
For Each I In Ay
    If Not Ay_XHas(A, I) Then Exit Function
Next
Ay_XHasAy = True
End Function

Function AyApHasEle(ParamArray AyAp()) As Boolean
Dim Av(): Av = AyAp
Dim Ay
For Each Ay In AyNz(Av)
    If Sz(Ay) > 0 Then AyApHasEle = True: Exit Function
Next
End Function

Function Ay_XHasAyChk(A, B) As String()
Dim C
C = AyMinus(B, A)
If Sz(C) = 0 Then Exit Function
XThw "Ay_XHasAyChk", "[Some-Ele] in [Ay-B] not [Ay-A]", "Some-Ele Ay-B Ay-A", C, B, A
End Function

Function Ay_XHasAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Sz(B) = 0 Then Stop
For Each BItm In B
    Ix = Ay_Ix_FmIx(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
Ay_XHasAyInSeq = True
End Function

Function Ay_XHasDupEle(A) As Boolean
If Sz(A) = 0 Then Exit Function
Dim Pool: Pool = A: Erase Pool
Dim I
For Each I In A
    If Ay_XHas(Pool, I) Then Ay_XHasDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function Ay_XHasNegOne(A) As Boolean
Dim V
If Sz(A) = 0 Then Exit Function
For Each V In A
    If V = -1 Then Ay_XHasNegOne = True: Exit Function
Next
End Function

Function Ay_XHasPredPXTrue(A, PX$, P) As Boolean
Dim X
For Each X In AyNz(A)
    If Run(PX, P, X) Then Ay_XHasPredPXTrue = True: Exit Function
Next
End Function

Function Ay_XHasPredXPTrue(A, XP$, P) As Boolean
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Run(XP, X, P) Then
        Ay_XHasPredXPTrue = True
        Exit Function
    End If
Next
End Function

Function Ay_XHasSubAy(A, SubAy) As Boolean
Const CSub$ = CMod & "Ay_XHasSubAy"
If Sz(A) = 0 Then Exit Function
If Sz(SubAy) = 0 Then XThw CSub, "Given {SubAy} is empty", ""
Dim I
For Each I In SubAy
    If Not Ay_XHas(A, I) Then Exit Function
Next
End Function

Private Sub ZZ_Ay_XHasAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert Ay_XHasAyInSeq(A, B) = True

End Sub

Private Sub ZZ_Ay_XHasDupEle()
Ass Ay_XHasDupEle(Array(1, 2, 3, 4)) = False
Ass Ay_XHasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub
