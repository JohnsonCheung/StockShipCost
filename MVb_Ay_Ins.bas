Attribute VB_Name = "MVb_Ay_Ins"
Option Compare Binary
Option Explicit

Function AyIns2(A, X1, X2, Optional At&)
Dim O
O = Ay_XReSz(A, At, 2)
Asg X1, O(At)
Asg X2, O(At + 1)
AyIns2 = O
End Function

Function AyIns(A, Optional M, Optional At = 0)
If 0 > At Or At > Sz(A) Then Stop
Dim O
O = Ay_XReSz(A, At)
If Not IsMissing(M) Then
    Asg M, O(At)
End If
AyIns = O
End Function
Private Sub Z_AyIns()
Dim A(), M, At&
'--
A = Array(1, 2, 3, 4, 5)
M = "a"
At = 2
Ept = Array(1, 2, "a", 3, 4, 5)
GoSub Tst
'
Exit Sub
Tst:
    Act = AyIns(A, M, At)
    C
    Return
End Sub
Function AyInsAy(A, B, Optional At&)
Dim O, NB&, J&
NB = Sz(B)
O = Ay_XReSz(A, At, NB)
For J = 0 To NB - 1
    Asg B(J), O(At + J)
Next
AyInsAy = O
End Function

Private Function Ay_XReSz(A, At, Optional Cnt = 1)
If Cnt < 1 Then Stop
Dim P1, P3
    P3 = AyMid(A, At)
    P1 = A
    If At = 0 Then
        Erase P1
        ReDim Preserve P1(Cnt - 1)
    Else
        ReDim Preserve P1(At + Cnt - 1)
    End If
Ay_XReSz = AyAp_XAdd(P1, P3)
End Function
Function AyEmpEle(A)
Dim O: O = A: Erase O
ReDim O(0)
AyEmpEle = O(0)
End Function

Private Sub Z_Ay_XReSz()
Dim Ay(), At&, Cnt&
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Ept = Array(1, Empty, Empty, Empty, 2, 3)
Exit Sub
Tst:
    Act = Ay_XReSz(Ay, At, Cnt)
    Ass Ay_IsEq(Act, Ept)
End Sub


Private Sub Z()
Z_AyIns
Z_Ay_XReSz
MVb_Ay_Ins:
End Sub
