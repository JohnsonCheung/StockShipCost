Attribute VB_Name = "MVb_Ay_Sub_Exl"
Option Compare Binary
Option Explicit
Function Ay_XExl_Patn(A, Patn$) As String()
Dim I, Re As New RegExp
Re.Pattern = Patn
For Each I In AyNz(A)
    If Not Re.Test(I) Then PushI Ay_XExl_Patn, I
Next
End Function
Function Ay_XExl_AtCnt(A, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Stop
If Sz(A) = 0 Then Ay_XExl_AtCnt = A: Exit Function
If At = 0 Then
    If Sz(A) = Cnt Then
        Ay_XExl_AtCnt = Ay_XCln(A)
        Exit Function
    End If
End If
Dim U&: U = UB(A)
If At > U Then Stop
If At < 0 Then Stop
Dim O: O = A
Dim J&
If IsObject(A(0)) Then
    For J = At To U - Cnt
        Set O(J) = O(J + Cnt)
    Next
Else
    For J = At To U - Cnt
        O(J) = O(J + Cnt)
    Next
End If
ReDim Preserve O(U - Cnt)
Ay_XExl_AtCnt = O
End Function

Function Ay_XExl_DDLin(A) As String()
Ay_XExl_DDLin = Ay_XWh_PredFalse(A, "Lin_IsDDLin")
End Function

Function Ay_XExl_DotLin(A) As String()
Ay_XExl_DotLin = Ay_XWh_PredFalse(A, "Lin_IsDotLin")
End Function

Function Ay_XExl_Ele(A, Ele)
Dim Ix&: Ix = Ay_Ix(A, Ele): If Ix = -1 Then Ay_XExl_Ele = A: Exit Function
Ay_XExl_Ele = Ay_XExl_EleAt(A, Ay_Ix(A, Ele))
End Function

Function Ay_XExl_EleAt(Ay, Optional At = 0, Optional Cnt = 1)
Ay_XExl_EleAt = Ay_XExl_AtCnt(Ay, At, Cnt)
End Function

Function Ay_XExl_EleLik(A, Lik$) As String()
If Sz(A) = 0 Then Exit Function
Dim J&
For J = 0 To UB(A)
    If A(J) Like Lik Then Ay_XExl_EleLik = Ay_XExl_EleAt(A, J): Exit Function
Next
End Function

Function Ay_XExl_EmpEle(A)
Dim O: O = Ay_XCln(A)
If Sz(A) > 0 Then
    Dim X
    For Each X In AyNz(A)
        PushNonEmp O, X
    Next
End If
Ay_XExl_EmpEle = O
End Function

Function Ay_XExl_EmpEleAtEnd(A)
Dim LasU&, U&
Dim O: O = A
For LasU = UB(A) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
Ay_XExl_EmpEleAtEnd = O
End Function

Function Ay_XExl_FmTo(A, FmIx, ToIx)
Dim U&
U = UB(A)
If 0 > FmIx Or FmIx > U Then XThw CSub, "[FmIx] is out of range", "Ay U FmIx ToIx", A, UB(A), FmIx, ToIx
If FmIx > ToIx Or ToIx > U Then XThw CSub, "[ToIx] is out of range", "Ay U FmIx ToIx", A, UB(A), FmIx, ToIx
Dim O
    O = A
    Dim I&, J&
    I = 0
    For J = ToIx + 1 To U
        O(FmIx + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = ToIx - FmIx + 1
    ReDim Preserve O(U - Cnt)
Ay_XExl_FmTo = O
End Function

Function Ay_XExl_FstEle(A)
Ay_XExl_FstEle = Ay_XExl_EleAt(A)
End Function

Function Ay_XExl_FstNEle(A, N)
Dim O: O = A
ReDim O(N - 1)
Dim J&
For J = 0 To UB(A) - N
    O(J) = A(N + J)
Next
Ay_XExl_FstNEle = O
End Function

Function Ay_XExl_FTIx(A, B As FTIx)
With B
    Ay_XExl_FTIx = Ay_XExl_FmTo(A, .FmIx, .ToIx)
End With
End Function

Function Ay_XExl_IxAy(A, IxAy)
'IxAy holds index if A to be remove.  It has been sorted else will be stop
Ass Ay_IsSrt(A)
Ass Ay_IsSrt(IxAy)
Dim J&
Dim O: O = A
For J = UB(IxAy) To 0 Step -1
    O = Ay_XExl_EleAt(O, CLng(IxAy(J)))
Next
Ay_XExl_IxAy = O
End Function

Function Ay_XExl_LasEle(A)
Ay_XExl_LasEle = Ay_XExl_EleAt(A, UB(A))
End Function

Function Ay_XExl_LasNEle(A, Optional NEle% = 1)
Dim O: O = A
Select Case Sz(A)
Case Is > NEle:    ReDim Preserve O(UB(A) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
Ay_XExl_LasNEle = O
End Function

Function Ay_XExl_Lik(A, Lik) As String()
Dim I
For Each I In AyNz(A)
    If Not I Like Lik Then PushI Ay_XExl_Lik, I
Next
End Function

Function Ay_XExl_LikAy(A, LikAy$()) As String()
Dim I
For Each I In AyNz(A)
    If Not IsInLikAy(I, LikAy) Then Push Ay_XExl_LikAy, I
Next
End Function

Function Ay_XExl_Likss(A, Likss$) As String()
Ay_XExl_Likss = Ay_XExl_LikAy(A, Ssl_Sy(Likss))
End Function

Function Ay_XExl_LikssAy(A, LikssAy$()) As String()
If Sz(LikssAy) = 0 Then Ay_XExl_LikssAy = Ay_Sy(A): Exit Function
Dim Likss
Stop
For Each Likss In AyNz(A)
    If Not IsInLikss(A, Likss) Then PushI Ay_XExl_LikssAy, A
Next
End Function

Function Ay_XExl_Neg(A)
Dim I
Ay_XExl_Neg = Ay_XCln(A)
For Each I In AyNz(A)
    If I >= 0 Then
        PushI Ay_XExl_Neg, I
    End If
Next
End Function

Function Ay_XExl_NEle(A, Ele, Cnt%)
If Cnt <= 0 Then Stop
Ay_XExl_NEle = Ay_XCln(A)
Dim X, C%
C = Cnt
For Each X In AyNz(A)
    If C = 0 Then
        PushI Ay_XExl_NEle, X
    Else
        If X <> Ele Then
            Push Ay_XExl_NEle, X
        Else
            C = C - 1
        End If
    End If
Next
X:
End Function

Function Ay_XExl_OneTermLin(A) As String()
Ay_XExl_OneTermLin = Ay_XWh_PredFalse(A, "Lin_IsOneTermLin")
End Function

Function Ay_XExl_Pfx(A, ExlPfx$) As String()
Dim I
For Each I In AyNz(A)
    If Not XHas_Pfx(I, ExlPfx) Then PushI Ay_XExl_Pfx, I
Next
End Function

Function Ay_XExl_T1Ay(A, ExlT1Ay0) As String()
'Exclude those Lin in Array-A its T1 in ExlT1Ay0
Dim Exl$(): Exl = CvNy(ExlT1Ay0): If Sz(Exl) = 0 Then Stop
Dim L
For Each L In AyNz(A)
    If Not Ay_XHas(Exl, Lin_T1(L)) Then
        PushI Ay_XExl_T1Ay, L
    End If
Next
End Function


Private Sub Z_Ay_XExl_AtCnt()
Dim A()
A = Array(1, 2, 3, 4, 5)
Ept = Array(1, 4, 5)
GoSub Tst
'
Exit Sub

Tst:
    Act = Ay_XExl_AtCnt(A, 1, 2)
    C
    Return
End Sub

Private Sub Z_Ay_XExl_EmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = Ay_XExl_EmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_Ay_XExl_FTIx()
Dim A
Dim FTIx1 As FTIx
Dim Act
A = SplitSpc("a b c d e")
Set FTIx1 = New_FTIx(1, 2)
Act = Ay_XExl_FTIx(A, FTIx1)
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_Ay_XExl_FTIx1()
Dim A
Dim Act
A = SplitSpc("a b c d e")
Act = Ay_XExl_FTIx(A, New_FTIx(1, 2))
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_Ay_XExl_IxAy()
Dim A(), IxAy
A = Array("a", "b", "c", "d", "e", "f")
IxAy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XExl_IxAy(A, IxAy)
    C
    Return
End Sub

Private Sub Z()
Z_Ay_XExl_AtCnt
Z_Ay_XExl_EmpEleAtEnd
Z_Ay_XExl_FTIx
Z_Ay_XExl_FTIx1
Z_Ay_XExl_IxAy
MVb_Ay_Sub_Exl:
End Sub
