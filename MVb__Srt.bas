Attribute VB_Name = "MVb__Srt"
Option Compare Binary
Option Explicit

Function LinesSrt$(A$)
LinesSrt = JnCrLf(Ay_XSrt(LinesSplit(A)))
End Function

Function Ay_IsSrt(A) As Boolean
Dim J&
For J = 0 To UB(A) - 1
   If A(J) > A(J + 1) Then Exit Function
Next
Ay_IsSrt = True
End Function

Function AyQSrt(A)
If Sz(A) = 0 Then Exit Function
Dim O: O = A
AyQSrtLH O, 0, UB(A)
AyQSrt = O
End Function

Sub AyQSrtLH(A, L&, H&)
If L >= H Then Exit Sub
Dim P&
P = AyQSrtPartition(A, L, H)
AyQSrtLH A, L, P
AyQSrtLH A, P + 1, H
End Sub

Function AyQSrtPartition&(A, L&, H&)
Dim V, I&, J&, X
V = A(L)
I = L - 1
J = H + 1
Dim Z&
Do
    Z = Z + 1
    If Z > 1000 Then Stop
    Do
        I = I + 1
    Loop Until A(I) >= V
    
    Do
        J = J - 1
    Loop Until A(J) <= V

    If I >= J Then
        AyQSrtPartition = J
        Exit Function
    End If

     X = A(I)
     A(I) = A(J)
     A(J) = X
Loop
End Function
Private Sub Z_Ay_XSrt__BY_AY()
Dim Ay, ByAy
Ay = Array(1, 2, 3, 4)
ByAy = Array(3, 4)
Ept = Array(3, 4, 1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XSrt__BY_AY(Ay, ByAy)
    C
    Return
End Sub

Function Ay_XSrt__BY_AY(Ay, ByAy)
Dim O: O = Ay_XCln(Ay)
Dim I
For Each I In ByAy
    If Ay_XHas(Ay, I) Then PushI O, I
Next
PushIAy O, AyMinus(Ay, O)
Ay_XSrt__BY_AY = O
End Function

Function Ay_XSrt(Ay, Optional Des As Boolean)
If Sz(Ay) = 0 Then Ay_XSrt = Ay: Exit Function
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = AyIns(O, Ay(J), Ay_XSrt__Ix(O, Ay(J), Des))
Next
Ay_XSrt = O
End Function

Private Function Ay_XSrt__Ix&(A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In A
        If V > I Then Ay_XSrt__Ix = O: Exit Function
        O = O + 1
    Next
    Ay_XSrt__Ix = O
    Exit Function
End If
For Each I In A
    If V < I Then Ay_XSrt__Ix = O: Exit Function
    O = O + 1
Next
Ay_XSrt__Ix = O
End Function

Function Ay_XSrt_IntoIxAy(Ay, Optional Des As Boolean) As Long()
If Sz(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyIns(O, J, Ay_XSrt_InToIxAy__Ix(O, Ay, Ay(J), Des))
Next
Ay_XSrt_IntoIxAy = O
End Function

Private Sub Z_Ay_XSrt()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                   Act = Ay_XSrt(A):       Ay_XAss_Eq Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = Ay_XSrt(A, True): Ay_XAss_Eq Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = Ay_XSrt(A):       Ay_XAss_Eq Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:XHas_Pfx:Function"
Push A, "Private:MdMthDrs_FunBdyLy:Function"
Push A, "Private:SrcMthFmIx_MthToIx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:XHas_Pfx:Function"
Push Exp, "Private:MdMthDrs_FunBdyLy:Function"
Push Exp, "Private:SrcMthFmIx_MthToIx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = Ay_XSrt(A)
Ay_XAss_Eq Exp, Act
End Sub

Private Function Ay_XSrt_InToIxAy__Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then Ay_XSrt_InToIxAy__Ix& = O: Exit Function
        O = O + 1
    Next
    Ay_XSrt_InToIxAy__Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then Ay_XSrt_InToIxAy__Ix& = O: Exit Function
    O = O + 1
Next
Ay_XSrt_InToIxAy__Ix& = O
End Function

Private Sub Z_Ay_XSrt_InToIxAy()
Dim A: A = Array("A", "B", "C", "D", "E")
Ay_XAss_Eq Array(0, 1, 2, 3, 4), Ay_XSrt_IntoIxAy(A)
Ay_XAss_Eq Array(4, 3, 2, 1, 0), Ay_XSrt_IntoIxAy(A, True)
End Sub

Private Function Ay_XSrt_InToIxAy_Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then Ay_XSrt_InToIxAy_Ix& = O: Exit Function
        O = O + 1
    Next
    Ay_XSrt_InToIxAy_Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then Ay_XSrt_InToIxAy_Ix& = O: Exit Function
    O = O + 1
Next
Ay_XSrt_InToIxAy_Ix& = O
End Function


Function Dic_XSrt(A As Dictionary) As Dictionary
If A.Count = 0 Then Set Dic_XSrt = New Dictionary: Exit Function
Dim K
Set Dic_XSrt = New Dictionary
For Each K In AyQSrt(A.Keys)
   Dic_XSrt.Add K, A(K)
Next
End Function

Private Sub ZZ_Ay_XSrt()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                    Act = Ay_XSrt(A):        Ay_IsEq_XAss Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = Ay_XSrt(A, True): Ay_IsEq_XAss Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = Ay_XSrt(A):       Ay_IsEq_XAss Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:XHas_Pfx:Function"
Push A, "Private:MdMthDrs_FunBdyLy:Function"
Push A, "Private:SrcMthFmIx_MthToIx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:XHas_Pfx:Function"
Push Exp, "Private:MdMthDrs_FunBdyLy:Function"
Push Exp, "Private:SrcMthFmIx_MthToIx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = Ay_XSrt(A)
Ay_IsEq_XAss Exp, Act
End Sub

Private Sub ZZ_Ay_XSrt_InToIxAy()
Dim A: A = Array("A", "B", "C", "D", "E")
Ay_IsEq_XAss Array(0, 1, 2, 3, 4), Ay_XSrt_IntoIxAy(A)
Ay_IsEq_XAss Array(4, 3, 2, 1, 0), Ay_XSrt_IntoIxAy(A, True)
End Sub


Private Sub Z()
Z_Ay_XSrt
Z_Ay_XSrt_InToIxAy
MVb__Srt:
End Sub
