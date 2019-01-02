Attribute VB_Name = "MTp_Tp_Lin_Cln"
Option Compare Binary
Option Explicit
Function ClnBrk1(A$(), Ny0) As Variant()
Dim O(), U%, Ny$(), L, T1$, T2$, NmDic As Dictionary, Ix%, Er$()
Ny = CvNy(Ny0)
U = UB(Ny)
ReDim O(U)
O = AyMap(O, "EmpSy")
Set NmDic = AyIxDic(Ny)
For Each L In A
    Lin_TRstAsg LTrim(L), T1, T2
    If NmDic.Exists(T1) Then
        Ix = NmDic(T1)
        Push O(Ix), T2 '<----
    End If
Next
Push O, ClnT1Chk(A, Ny)
ClnBrk1 = O
End Function

Function ClnT1Chk(A$(), T1Ay0) As String()
Dim T1Ay$(), L, O$()
T1Ay = CvNy(T1Ay0)
For Each L In A
    If Not Ay_XHas(T1Ay, Lin_T1(L)) Then Push O, L
Next
If Sz(O) > 0 Then
    O = Ay_XAdd_Pfx(Ay_XQuote_SqBkt(O), Space(4))
    O = AyIns(O, QQ_Fmt("Following lines have invalid T1.  Valid T1 are [?]", JnSpc(T1Ay)))
End If
ClnT1Chk = O
End Function

Function LinCln$(A)
If IsEmp(A) Then Exit Function
If Lin_IsDotLin(A) Then Exit Function
If Lin_IsSngTerm(A) Then Exit Function
If Lin_IsDDLin(A) Then Exit Function
LinCln = XTak_BefDD(A)
End Function

Function LyCln(A) As String()
LyCln = Ay_XExl_EmpEle(AyMap_Sy(A, "LinCln"))
End Function

Function LyClnLnxAy(A) As Lnx()
Dim O()  As Lnx, L$, J%
For J = 0 To UB(A)
    L = LinCln(A(J))
    If L <> "" Then
        Dim M  As Lnx
        Set M = New Lnx
        M.Ix = J
        M.Lin = A(J)
        Push O, M
    End If
Next
LyClnLnxAy = O
End Function
