Attribute VB_Name = "MVb_Dic_Rel"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Dic_Rel."

Function CvRel(A) As Rel
Set CvRel = A
End Function

Property Get EmpRel() As Rel
Set EmpRel = New Rel
Set EmpRel.Rel = New Dictionary
End Property

Function IsRel(A) As Boolean
IsRel = TypeName(A) = "Rel"
End Function

Function NewRel(Ly$()) As Rel
Dim O As New Rel, K
Set O.Rel = New_Dic_LY(Ly)
For Each K In O.Rel.Keys
    Set O.Rel(K) = New_ASet_SSL(O.Rel(K))
Next
Set NewRel = O
End Function

Function NewRelVBL(Vbl$) As Rel
Set NewRelVBL = NewRel(VblLy(Vbl))
End Function

Function RelClone(A As Rel) As Rel
Set RelClone = New Rel
Set RelClone.Rel = DicClone(A.Rel)
End Function

Function RelCnt&(A As Rel)
RelCnt = A.Rel.Count
End Function

Sub RelDmp(A As Rel)
Ay_XDmp RelFmt(A)
End Sub

Function RelFmt(A As Rel) As String()
Dim K
For Each K In A.Rel.Keys
    PushI RelFmt, RelParLin(A, K)
Next
End Function

Function RelHasPar(A As Rel, Par) As Boolean
RelHasPar = A.Rel.Exists(Par)
End Function

Function RelIsEq(A As Rel, B As Rel) As Boolean
If Not ASet_IsEq(A.Rel, B.Rel) Then Exit Function
Dim K
For Each K In RelKeys(A)
    If Not ASet_IsEq(A.Rel(K), B.Rel(K)) Then Exit Function
Next
RelIsEq = True
End Function

Sub RelIsEq_XAss(A As Rel, B As Rel, Optional Msg$ = "Two rel are diff", Optional ANm$ = "Rel-A", Optional BNm$ = "Rel-B")
If RelIsEq(A, B) Then Exit Sub
Dim O$()
PushI O, Msg
PushI O, QQ_Fmt("?-ParCnt(?) / ?-ParCnt(?)", ANm, RelParCnt(A), BNm, RelParCnt(B))
PushI O, ANm & " --------------------"
PushIAy O, RelFmt(A)
PushI O, BNm & " --------------------"
PushIAy O, RelFmt(B)
Ay_XBrw_XHalt O
End Sub

Sub RelIsVdtAss(A As Rel)
Const CSub$ = CMod & "RelIsVdtAss"
Dim K
For Each K In A.Rel.Keys
    If Not IsASet(A.Rel(K)) Then
        XThw CSub, "Given Rel is not a valid due to the chd of K is not ASet", "Rel K [TypeName of K's Chd]", RelFmt(A), K, TypeName(A.Rel(K))
    End If
Next
End Sub

Function RelItmCnt&(A As Rel)
RelItmCnt = ASet_Cnt(RelItms(A))
End Function

Function RelItmIsLeaf(A As Rel, Itm) As Boolean
RelItmIsLeaf = Not RelItmIsPar(A, Itm) Or RelItmIsNoChdPar(A, Itm)
End Function

Function RelItmIsNoChdPar(A As Rel, Itm) As Boolean
If Not RelItmIsPar(A, Itm) Then Exit Function
RelItmIsNoChdPar = ASet_Cnt(CvSet(A.Rel(Itm))) = 0
End Function

Function RelItmIsPar(A As Rel, Itm) As Boolean
RelItmIsPar = A.Rel.Exists(Itm)
End Function

Function RelItms(A As Rel) As ASet
Dim O As ASet, K
Set O = EmpASet
ASet_XAdd_Itr O, A.Rel.Keys
For Each K In A.Rel.Keys
    ASet_Push O, RelParChd(A, K)
Next
Set RelItms = O
End Function

Function RelItms_DPD_ORD(A As Rel) As ASet
'Return itms in Rel in dependant order. Throw er if there is cyclic
'Example: A B C D
'         C D E
'         E X
'Return: B D X E C A
Const CSub$ = CMod & "RelItms_DPD_ORD"
Dim O As ASet, J%, M As Rel, Leaves As ASet
Set O = EmpASet
Set M = RelClone(A)
Do
    J = J + 1: If J > 1000 Then XThw_Msg CSub, "looping to much"
    Set Leaves = RelLeaves(M)
    If ASet_Cnt(Leaves) = 0 Then
        If RelCnt(M) > 0 Then
            XThw CSub, "Cyclic relation is found so far.  No leaves but there is remaining Rel", _
            "Turn-Cnt [Orginal rel] [Dpd itm found] [Remaining relation not solved]", _
            J, RelFmt(A), ASet_Lin(O), RelFmt(M)
        End If
        Set RelItms_DPD_ORD = O
        Exit Function
    End If
    ASet_Push O, Leaves
    RelRmvLeaf M
    ASet_Push O, RelNoChdPar(M)
    RelRmvNoChdPar M
Loop
RelItms_DPD_ORD = O
End Function

Function RelKeys(A As Rel)
RelKeys = A.Rel.Keys
End Function

Function RelLeaves(A As Rel) As ASet
Dim Itm, O As ASet
Set O = EmpASet
For Each Itm In ASet_Itms(RelItms(A))
    If RelItmIsLeaf(A, Itm) Then ASet_XPush O, Itm
Next
Set RelLeaves = O
End Function

Function RelNoChdPar(A As Rel) As ASet
Dim O As ASet, Par
Set O = EmpASet
For Each Par In A.Rel.Keys
    If RelParIsNoChd(A, Par) Then ASet_XPush O, Par
Next
Set RelNoChdPar = O
End Function

Function RelPar(A As Rel) As ASet
RelPar = New_ASet(A.Rel.Keys)
End Function

Sub RelParAss(A As Rel, Par, Fun$)
If A.Rel.Exists(Par) Then Exit Sub
XThw Fun, "Given Par is not a parent", "Rel Par", RelFmt(A), Par: Stop
End Sub

Function RelParChd(A As Rel, Par) As ASet
Const CSub$ = CMod & "RelParChd"
RelParAss A, Par, CSub
Set RelParChd = A.Rel(Par)
End Function

Function RelParCnt&(A As Rel)
RelParCnt = A.Rel.Count
End Function

Function RelParHasChd(A As Rel, Par, Chd) As Boolean
RelParHasChd = ASet_XHas(RelParChd(A, Par), Chd)
End Function

Function RelParIsNoChd(A As Rel, Par) As Boolean
Const CSub$ = CMod & "RelParIsNoChd"
If A.Rel.Exists(Par) Then
    RelParIsNoChd = ASet_Cnt(A.Rel(Par)) = 0: Exit Function
End If
XThw CSub, "Given Par is not a par", "Par Rel", Par, RelFmt(A)
End Function

Function RelParLin$(A As Rel, K)
Const CSub$ = CMod & "RelParLin"
RelParAss A, K, CSub
Dim X
Asg A.Rel(K), X
If IsASet(X) Then
    RelParLin = K & " " & ASet_Lin(A.Rel(K))
Else
    RelParLin = K & " [*Chd is not ASet, But *" & TypeName(X) & "]"
End If
End Function

Function RelParRmvChd(A As Rel, Par, Chd) As Boolean
Dim X As ASet
If RelParHasChd(A, Par, Chd) Then
    Set X = RelParChd(A, Par)
    ASet_XRmv_Itm X, Chd
    Set A.Rel(Par) = X
    RelParRmvChd = True
End If
End Function

Function RelRmvLeaf&(A As Rel)
Dim Par, Leaf, O&
For Each Leaf In ASet_Itms(RelLeaves(A))
    For Each Par In A.Rel.Keys
        If RelParRmvChd(A, Par, Leaf) Then
            O = O + 1
        End If
    Next
Next
RelRmvLeaf = O
End Function

Sub RelRmvNoChdPar(O As Rel)
Dim NoChdPar
For Each NoChdPar In ASet_Itms(RelNoChdPar(O))
    RelRmvPar O, NoChdPar
Next
End Sub

Sub RelRmvPar(O As Rel, Par)
Const CSub$ = CMod & "RelRmvPar"
If RelHasPar(O, Par) Then
    O.Rel.Remove Par
    Exit Sub
End If
XThw CSub, "Given Par is not a par", "Par Rel", Par, RelFmt(O)
End Sub

Property Get SampRel() As Rel
Set SampRel = NewRelVBL("B C D | D E | X")
End Property

Private Sub Z_RelItms()
Dim Act As ASet, Ept As ASet, A As Rel
Set Ept = New_ASet_SSL("A B C D E")
Set A = NewRelVBL("A B C | B D E | C D")
GoSub Tst
Exit Sub
Tst:
    Set Act = RelItms(A)
    C
    Return
End Sub

Private Sub Z_RelItms_DPD_ORD()
Dim Act As ASet, Ept As ASet
Dim Rel As Rel
GoSub T1
GoSub T2
Exit Sub
T1:
    Set Ept = New_ASet_SSL("C E X D B")
    Set Rel = NewRelVBL("B C D | D E | X")
    GoSub Tst
    Return
'
T2:
    Dim X$()
    PushI X, "MVb"
    PushI X, "MIde MVb MXls MAcs"
    PushI X, "MXls MVb"
    PushI X, "MDao MVb MDta"
    PushI X, "MAdo MVb"
    PushI X, "MAdoX MVb"
    PushI X, "MApp  MVb"
    PushI X, "MDta  MVb"
    PushI X, "MTp   MVb"
    PushI X, "MSql  MVb"
    PushI X, "AStkShpCst MVb MXls MAcs"
    PushI X, "MAcs  MVb MXls"
    Set Rel = NewRel(X)
    Set Ept = New_ASet_SSL("MVb MIde MXls MDao MAdo MAdoX MApp MDta MTp MSql AStkShpCst MAcs ")
    GoSub Tst
    Return
Tst:
    Set Act = RelItms_DPD_ORD(Rel)
    ASet_IsEq_XAss Act, Ept
    Return
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$()
Dim C$
Dim D As Rel

CvRel A
IsRel A
NewRel B
NewRelVBL C
RelClone D
RelCnt D
RelDmp D
RelFmt D
RelHasPar D, A
RelIsEq D, D
RelIsEq_XAss D, D, C, C, C
RelIsVdtAss D
RelItmCnt D
RelItmIsLeaf D, A
RelItmIsNoChdPar D, A
RelItmIsPar D, A
RelItms D
RelItms_DPD_ORD D
RelKeys D
RelLeaves D
RelNoChdPar D
RelPar D
RelParAss D, A, C
RelParChd D, A
RelParCnt D
RelParHasChd D, A, A
RelParIsNoChd D, A
RelParLin D, A
RelParRmvChd D, A, A
RelRmvLeaf D
RelRmvNoChdPar D
RelRmvPar D, A
End Sub

Private Sub Z()
Z_RelItms
Z_RelItms_DPD_ORD
End Sub
