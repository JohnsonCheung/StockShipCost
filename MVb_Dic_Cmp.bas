Attribute VB_Name = "MVb_Dic_Cmp"
Option Compare Binary
Option Explicit
Private Type DicCmp
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Private A As Dictionary, B As Dictionary
Private A_Nm$, B_Nm$

Function CmpRslt_Fmt(A As DicCmp) As String()
With A
    CmpRslt_Fmt = AyAp_XAdd( _
        WFmt_Excess(.AExcess, A_Nm), _
        WFmt_Excess(.BExcess, B_Nm), _
        WFmt_Dif(.ADif, .BDif), _
        WFmt_Sam(.Sam))
End With
End Function

Function Dic_Cmp(A_Dic As Dictionary, B_Dic As Dictionary, ANm$, BNm$) As DicCmp
Set A = A_Dic
Set B = B_Dic
A_Nm = ANm
B_Nm = BNm
With Dic_Cmp
    Set .AExcess = Dic_XMinus(A, B)
    Set .BExcess = Dic_XMinus(B, A)
    Set .Sam = Dic_XUnion(A, B)
    Dic_XUnionKey .ADif, .BDif
End With
End Function

Function Dic_CmpFmt(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
Dic_CmpFmt = CmpRslt_Fmt(Dic_Cmp(A, B, Nm1, Nm2))
End Function

Sub Dic_Cmp_XBrw(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
Ay_XBrw Dic_CmpFmt(A, B, Nm1, Nm2)
End Sub

Function Dic_XUnion(A As Dictionary, B As Dictionary) As Dictionary
Set Dic_XUnion = New Dictionary
If A.Count = 0 Or B.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            Dic_XUnion.Add K, A(K)
        End If
    End If
Next
End Function

Private Sub Dic_XUnionKey( _
    OADif As Dictionary, OBDif As Dictionary)
Dim K
Set OADif = New Dictionary
Set OBDif = New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            OADif.Add K, A(K)
            OBDif.Add K, B(K)
        End If
    End If
Next
End Sub

Private Function WFmt_Dif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Stop
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S(0) As S1S2, Ly$(), KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(KK) & vbCrLf & B(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2Ay_Fmt(S)
    PushAy O, Ly
Next
WFmt_Dif = O
End Function

Private Function WFmt_Excess(A As Dictionary, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S(0) As S1S2
S2 = "!" & "Er Excess (" & Nm & ")"
For Each K In A.Keys
    S1 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    PushAy WFmt_Excess, S1S2Ay_Fmt(S)
Next
End Function

Private Function WFmt_Sam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S1S2, KK$
For Each K In A.Keys
    KK = K
    PushObj S, S1S2("*Same", K & vbCrLf & LinesUnderLin(KK) & vbCrLf & A(K))
Next
WFmt_Sam = S1S2Ay_Fmt(S)
End Function

Private Sub Z_Dic_Cmp_XBrw()
Set A = NewDicVBL("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
Set B = NewDicVBL("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
Dic_Cmp_XBrw A, B
End Sub

Private Sub Z()
Z_Dic_Cmp_XBrw
End Sub
