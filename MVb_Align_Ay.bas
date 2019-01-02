Attribute VB_Name = "MVb_Align_Ay"
Option Compare Binary
Option Explicit

Function Ay_XAlign_NTerm(A, N%) As String()
Dim W%(), L
W = ZWdtAy(A, N)
For Each L In AyNz(A)
    PushI Ay_XAlign_NTerm, Ay_XAlign_NTerm1(L, W)
Next
End Function

Function Ay_XAlign_T1(A) As String()
Dim T1$(), Rest$()
    AyAsgT1AyRestAy A, T1, Rest
T1 = Ay_XAlign_L(T1)
Ay_XAlign_T1 = AyAB_XJn(T1, Rest)
End Function

Private Function Ay_XAlign_NTerm1$(A, W%())
Dim Ay$(), J%, N%, O$(), I
N = Sz(W)
Ay = LinN_NTermRst(A, N)
If Sz(Ay) <> N + 1 Then Stop
For J = 0 To N - 1
    PushI O, XAlignL(Ay(J), W(J))
Next
PushI O, Ay(N)
Ay_XAlign_NTerm1 = RTrim(JnSpc(O))
End Function

Private Function ZWdtAy(A, NTerm%) As Integer()
If Sz(A) = 0 Then Exit Function
Dim O%(), W%(), L
ReDim O(NTerm - 1)
For Each L In A
    W = ZWdtAy1(L, NTerm)
    O = ZWdtAy2(O, W)
Next
ZWdtAy = O
End Function
Private Function ZWdtAy1(Lin, N%) As Integer()
Dim T
For Each T In LinNTerm(Lin, N)
    PushI ZWdtAy1, Len(T)
Next
End Function
Private Function ZWdtAy2(A%(), B%()) As Integer()
Dim O%(), J%, I
O = A
For Each I In B
    If I > O(J) Then O(J) = I
    J = J + 1
Next
ZWdtAy2 = O
End Function

Function Ay_XAlign_AtChr(A, AtChr$) As String()
Dim T1$(), Rst$(), I, P%
For Each I In AyNz(A)
    P = InStr(I, AtChr)
    If P = 0 Then
        PushI T1, ""
        PushI Rst, I
    Else
        PushI T1, Left(I, P)
        PushI Rst, Mid(I, P + 1)
    End If
Next
Dim J&
T1 = Ay_XAlign_R(T1)
For Each I In AyNz(T1)
    PushI Ay_XAlign_AtChr, I & Rst(J)
    J = J + 1
Next
End Function

Function Ay_XAlign_AtDot(A) As String()
Ay_XAlign_AtDot = Ay_XAlign_AtChr(A, ".")
End Function

Function Ay_XAlign_1T(A) As String()
Ay_XAlign_1T = Ay_XAlign_NTerm(A, 1)
End Function

Function Ay_XAlign_2T(A) As String()
Ay_XAlign_2T = Ay_XAlign_NTerm(A, 2)
End Function

Function Ay_XAlign_3T(A$()) As String()
Ay_XAlign_3T = Ay_XAlign_NTerm(A, 3)
End Function

Function Ay_XAlign_4T(A$()) As String()
Ay_XAlign_4T = Ay_XAlign_NTerm(A, 4)
End Function

Function Ay_XAlign_L(Ay) As String()
Dim W%: W = Ay_Wdt(Ay) + 1
Dim I
For Each I In AyNz(Ay)
    Push Ay_XAlign_L, XAlignL(I, W)
Next
End Function

Function Ay_XAlign_R(Ay) As String()
Dim W%: W = Ay_Wdt(Ay)
Dim I
For Each I In AyNz(Ay)
    Push Ay_XAlign_R, XAlignR(I, W)
Next
End Function

Private Sub Z_Ay_XAlign_2T()
Dim Ly$()
Ly = Ap_Sy("AAA B C D", "A BBB CCC")
Ept = Ap_Sy("AAA B   C D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XAlign_2T(Ly)
    C
    Return
End Sub
Private Sub Z_Ay_XAlign_3T()
Dim Ly$()
Ly = Ap_Sy("AAA B C D", "A BBB CCC")
Ept = Ap_Sy("AAA B   C   D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = Ay_XAlign_3T(Ly)
    C
    Return
End Sub




Private Sub Z()
Z_Ay_XAlign_2T
Z_Ay_XAlign_3T
MVb_Align_Ay:
End Sub
