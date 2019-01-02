Attribute VB_Name = "MVb_Ay_IxAy"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Ay_IxAy."

Function Ay_Ix_FmIx&(A, Itm, Fm&)
Dim O&
For O = Fm To UB(A)
    If A(O) = Itm Then Ay_Ix_FmIx = O: Exit Function
Next
Ay_Ix_FmIx = -1
End Function

Function Ay_LinPair_FstLin_IsCnt_SndLin_IsItm(A) As String()
'It is 2 line first line is 0 ...
'first line is x0 x1 of A$()
Dim U&: U = UB(A)
If U = -1 Then Exit Function
Dim A1$()
Dim A2$()
ReSz A1, U
ReSz A2, U
Dim O$(), J%, L$, W%
For J = 0 To U
    L = Len(A(J))
    W = Max(L, Len(J))
    A1(J) = XAlignL(J, W)
    A2(J) = XAlignL(A(J), W)
Next
Ay_LinPair_FstLin_IsCnt_SndLin_IsItm = Ap_Sy(JnSpc(A1), JnSpc(A2))
End Function

Function Ay_Ix&(A, M)
Dim J&
For J = 0 To UB(A)
    If A(J) = M Then Ay_Ix = J: Exit Function
Next
Ay_Ix = -1
End Function

Function Ay_IxAy(A, SubAy, Optional ChkNotFound As Boolean) As Long()
Dim I
For Each I In AyNz(SubAy)
    PushI Ay_IxAy, Ay_Ix(A, I)
Next
End Function

Sub Ay_IxAyAsg(Dr, IxAy%(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    If IsObject(OAp(J)) Then
        Set OAp(J) = Dr(IxAy(J))
    Else
        OAp(J) = Dr(IxAy(J))
    End If
Next
End Sub

Sub Ay_IxAyAsgAp(A, IxAy&(), ParamArray OAp())
Dim J&
For J = 0 To UB(IxAy)
    Asg A(IxAy(J)), OAp(J)
Next
End Sub

Function Ay_IxAyI(A, B) As Integer()
Ay_IxAyI = Ay_IxAyInto(A, B, EmpIntAy)
End Function

Function Ay_IxAyInto(A, B, OIntoAy)
Dim J&, U&, O
O = OIntoAy
Erase O
U = UB(B)
ReDim O(U)
For J = 0 To U
    O(J) = Ay_Ix(A, B(J))
Next
Ay_IxAyInto = O
End Function
Function Ay_NegEleDic(A) As Dictionary
Dim O As New Dictionary
Dim J&, I
For Each I In AyNz(A)
    If I < 0 Then O.Add J, I
    J = J + 1
Next
Set Ay_NegEleDic = O
End Function

Sub Ay_XAss_NegEle(A)
Dim O As Dictionary
Set O = Ay_NegEleDic(A)
If O.Count = 0 Then Exit Sub
XThw CSub, "Neg element found in Ay", "NegEle-Ix-Val Ay", O, A
End Sub

Function Ay_IsSamSz(A, B) As Boolean
Ay_IsSamSz = Sz(A) = Sz(B)
End Function

Sub Ay_XAss_SamSz(A, B, Fun$, N1$, N2$)
If Ay_IsSamSz(A, B) Then Exit Sub
XThw Fun, "Two array are different size", "Sz1 Sz2 Ay1Nm Ay2Nm Ay1 Ay2", Sz(A), Sz(B), N1, N2, A, B
End Sub
Function Ay_XAdd(A, B)
Dim O
O = A
PushIAy O, B
Ay_XAdd = O
Stop
End Function

Sub Ay_XAss_Dup(A, Fun$)
Const CSub$ = CMod & "Ay_XAss_Dup"
Dim Dup
Dup = Ay_XWh_Dup(A)
If Sz(Dup) = 0 Then Exit Sub
XThw CSub, "There is dup in given Ay", "Dup Ay", Dup, A
End Sub

Function U_IntAy(U&) As Integer()
Dim J&
For J = 0 To U
    PushI U_IntAy, J
Next
End Function

Function U_IxAy(U&) As Long()
Dim O&()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = J
    Next
U_IxAy = O
End Function

Private Sub ZZ()
Dim A&
Dim B
Dim C$
Dim D As Boolean
Dim E%()
Dim F()
Dim G&()
Dim XX
Ay_Ix_FmIx B, B, A
Ay_LinPair_FstLin_IsCnt_SndLin_IsItm B
Ay_Ix B, B
Ay_IxAy B, B, D
Ay_IxAyAsg B, E, F
Ay_IxAyAsgAp B, G, F
Ay_IxAyI B, B
Ay_IxAyInto B, B, B
Ay_XAdd B, B
Ay_XAss_Dup B, C
U_IntAy A
U_IxAy A
End Sub

Private Sub Z()
End Sub
