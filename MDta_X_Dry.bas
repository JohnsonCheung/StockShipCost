Attribute VB_Name = "MDta_X_Dry"
Option Compare Binary
Option Explicit
Const CMod$ = "MDta_X_Dry."
Function AyConst_ValConstDry(A, Constant) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I
For Each I In A
   Push O, Array(I, Constant)
Next
AyConst_ValConstDry = O
End Function

Function Dry_XWh_ColInAy(A(), ColIx%, InAy) As Variant()
Const CSub$ = CMod & "Dry_XWh_ColInAy"
If Not IsArray(InAy) Then XThw CSub, "[InAy] is not Array, but [TypeName]", InAy, TypeName(InAy)
If Sz(InAy) = 0 Then Dry_XWh_ColInAy = A: Exit Function
Dim Dr
For Each Dr In AyNz(A)
    If Ay_XHas(InAy, Dr(ColIx)) Then PushI Dry_XWh_ColInAy, Dr
Next
End Function
Sub C3DryDo(C3Dry(), ABC$)
If Sz(C3Dry) = 0 Then Exit Sub
Dim Dr
For Each Dr In C3Dry
    Run ABC, Dr(0), Dr(1), Dr(2)
Next
End Sub

Sub C4DryDo(C4Dry(), ABCD$)
If Sz(C4Dry) = 0 Then Exit Sub
Dim Dr
For Each Dr In C4Dry
    Run ABCD, Dr(0), Dr(1), Dr(2), Dr(3)
Next
End Sub

Function DotNyDry(DotNy$()) As Variant()
If Sz(DotNy) = 0 Then Exit Function
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, Ap_Sy(.S1, .S2)
   End With
Next
DotNyDry = O
End Function
Private Sub ZZ_Dry_Fmt()
Ay_XDmp Dry_Fmt(Samp_Dry1)
End Sub


Function Dry_XWh_ColHasDup(A(), ColIx%) As Variant()
Dim B(): B = Ay_XWh_Dup(DryCol(A, ColIx))
Dry_XWh_ColHasDup = Dry_XWh_ColInAy(A, ColIx, B)
End Function

Function DryFstColEqV(A(), ColIx%, V)
Dim Dr
For Each Dr In AyNz(A)
    If Dr(ColIx) = V Then DryFstColEqV = Dr: Exit Function
Next
End Function

Private Function Dry_MgeIx&(Dry(), Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   Dry_MgeIx = O
   Exit Function
Nxt:
Next
Dry_MgeIx = -1
End Function
Function Dr_XAdd_CC(ODr, C1, C2, U%) As Variant()
If UB(ODr) + 2 > U Then Stop
ReDim Preserve ODr(U)
ODr(U - 1) = C1
ODr(U) = C2
Dr_XAdd_CC = ODr
End Function
Function Dry_XIns_CC(A(), C1, C2) As Variant()
Dim R&, Dr, O()
O = Ay_XReSz(O, A)
For Each Dr In AyNz(A)
    O(R) = AyIns2(Dr, C1, C2)
    R = R + 1
Next
Dry_XIns_CC = O
End Function
Function Dry_XAdd_CC(A(), C1, C2) As Variant()
Dim UCol%, R&, Dr, O()
O = Ay_XReSz(O, A)
UCol = Dry_NCol(A) + 1
For Each Dr In AyNz(A)
    O(R) = Dr_XAdd_CC(Dr, C1, C2, UCol)
    R = R + 1
Next
Dry_XAdd_CC = O
End Function

Function Dry_XAdd_Col(A(), C) As Variant()
Dim UCol%, R&, Dr, O()
O = Ay_XReSz(O, A)
UCol = Dry_NCol(A)
For Each Dr In AyNz(A)
    ReDim Preserve Dr(UCol)
    Dr(UCol) = C
    O(R) = Dr
    R = R + 1
Next
Dry_XAdd_Col = O
End Function

Function Dry_XAdd_ConstCol(Dry(), ConstVal) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim N%
   N = Sz(Dry(0))
Dim O()
   Dim Dr, J&
   ReDim O(UB(Dry))
   For Each Dr In Dry
       ReDim Preserve Dr(N)
       Dr(N) = ConstVal
       O(J) = Dr
       J = J + 1
   Next
Dry_XAdd_ConstCol = O
End Function

Function Dry_XAdd_ValIdCntCol(A, ColIx) As Variant() ' Add 2 col at end (Id and Cnt) according to col(ColIx)
Dim O(), NCol%, Dr, R&, D As Dictionary, UCol%, IdCnt&()
O = A
UCol = Dry_NCol(O) + 1   ' The UCol after add
Set D = DryColSeqCntDic(A, ColIx)
For Each Dr In A
    ReDim Preserve Dr(UCol)
    If Not D.Exists(Dr(ColIx)) Then Stop
    IdCnt = D(ColIx)
    Dr(UCol - 1) = IdCnt(0)
    Dr(UCol) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
Dry_XAdd_ValIdCntCol = O
End Function

Function Dry_XAdd_ValIdCol(A(), ValCol) As Variant()
Dim NCol%, Dic As Dictionary, O(), Dr, IdCnt, R&
NCol = Dry_NCol(A)
Set Dic = AyDistIdCntDic(DryCol(A, ValCol))
O = Ay_XReSz(O, A)
For Each Dr In AyNz(A)
    ReDim Preserve Dr(NCol + 1)
    IdCnt = Dic(Dr(ValCol))
    Dr(NCol) = IdCnt(0)
    Dr(NCol + 1) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
Dry_XAdd_ValIdCol = O
End Function

Sub Dry_XBrw(A, Optional MaxColWdt% = 100, Optional BrkColIx% = -1)
Ay_XBrw Dry_Fmt(A, MaxColWdt, BrkColIx)
End Sub

Function DryCntDic(A, KeyColIx%) As Dictionary
Dim O As New Dictionary
Dim J%, Dr, K
For J = 0 To UB(A)
    Dr = A(J)
    K = Dr(KeyColIx)
    If O.Exists(K) Then
        O(K) = O(K) + 1
    Else
        O.Add K, 1
    End If
Next
Set DryCntDic = O
End Function

Function DryCol(A, ColIx) As Variant()
DryCol = DryColInto(A, ColIx, Array())
End Function

Function DryColInto(A, ColIx, OInto)
Dim O, J&, Dr, U&
O = Ay_XReSz(OInto, A)
For Each Dr In AyNz(A)
    If UB(Dr) >= ColIx Then
        O(J) = Dr(ColIx)
    End If
    J = J + 1
Next
DryColInto = O
End Function

Function DryColSeqCntDic(A, ColIx) As Dictionary
Set DryColSeqCntDic = Ay_SeqCntDic(DryCol(A, ColIx))
End Function

Function DryColCntDic(A, ColIx) As Dictionary
Set DryColCntDic = AyCntDic(DryCol(A, ColIx))
End Function

Function DryColSqlTy$(A(), ColIx%)
Dim O As VbVarType, Dr, V, T As VbVarType
For Each Dr In A
    If UB(Dr) >= ColIx Then
        V = Dr(ColIx)
        T = VarType(V)
        If T = vbString Then
            If Len(V) > 255 Then DryColSqlTy = "Memo": Exit Function
        End If
        O = MaxVbTy(O, T)
    End If
Next
DryColSqlTy = VbTySqlTy(O)
End Function
Function VbTySqlTy$(A As VbVarType, Optional IsMem As Boolean)
Select Case A
Case vbEmpty: VbTySqlTy = "Text(255)"
Case vbBoolean: VbTySqlTy = "YesNo"
Case vbByte: VbTySqlTy = "Byte"
Case vbInteger: VbTySqlTy = "Short"
Case vbLong: VbTySqlTy = "Long"
Case vbDouble: VbTySqlTy = "Double"
Case vbSingle: VbTySqlTy = "Single"
Case vbCurrency: VbTySqlTy = "Currency"
Case vbDate: VbTySqlTy = "Date"
Case vbString: VbTySqlTy = IIf(IsMem, "Memo", "Text(255)")
Case Else: Stop
End Select
End Function

Sub Dry_XDmp_(A())
Ay_XDmp Dry_Fmtss(A)
End Sub

Sub Dry_XDmp_1(A())
Dry_FmtssDmp A
End Sub
Sub Dry_FmtssDmp(A())
D Dry_Fmtss(A)
End Sub


Function DryMAp_JnDot(A()) As String()
Dim Dr
For Each Dr In AyNz(A)
    PushI DryMAp_JnDot, JnDot(Dr)
Next
End Function
Function Dry_XIns_CxAp(A, ParamArray CxAp()) As Variant()
Dim Av(): Av = CxAp
Dry_XIns_CxAp = Dry_XIns_CxAv(A, Av)
End Function

Function Dry_XIns_CxAv(A, CxAv) As Variant()
'Called by DryInCC
Dim Dr, O()
If Sz(A) = 0 Then Exit Function
For Each Dr In A
    Push O, AyInsAy(Dr, CxAv)
Next
Dry_XIns_CxAv = O
End Function
Function Dry_XIns_C4(A, C1, C2, C3, C4) As Variant()
Dry_XIns_C4 = Dry_XIns_CxAv(A, Array(C1, C2, C3, C4))
End Function

Function Dry_XIns_CC1(A, C1, C2) As Variant()
Dry_XIns_CC1 = Dry_XIns_CxAv(A, Array(C1, C2))
End Function

Function Dry_XIns_CCC(A, C1, C2, C3) As Variant()
Dry_XIns_CCC = Dry_XIns_CxAv(A, Array(C1, C2, C3))
End Function

Function Dry_XIns_Col(A, C, Optional Ix&) As Variant()
Dim Dr
For Each Dr In A
    PushI Dry_XIns_Col, AyIns(Dr, C, At:=Ix)
Next
End Function

Function Dry_XIns_Const(A, C, Optional At& = 0) As Variant()
Dim O(), Dr
If Sz(A) = 0 Then Exit Function
For Each Dr In A
    Push O, AyIns(Dr, C, At)
Next
Dry_XIns_Const = O
End Function

Function DryIntCol(A, ColIx%) As Integer()
DryIntCol = DryColInto(A, ColIx, EmpIntAy)
End Function

Function DryIsBrkAtDrIx(Dry, DrIx&, BrkColIx%) As Boolean
If Sz(Dry) = 0 Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(Dry) Then Exit Function
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Function
DryIsBrkAtDrIx = True
End Function

Function DryKeyGpAy(Dry(), K_Ix%, Gp_Ix%) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In Dry
    K = Dr(K_Ix)
    Gp = Dr(Gp_Ix)
    O_Ix = Ay_Ix(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryKeyGpAy = O
End Function


Function DryMge(Dry, MgeIx%, Sep$) As Variant()
Dim O(), J%
Dim Ix%
For J = 0 To UB(Dry)
   Ix = DryMgeIx(O, Dry(J), MgeIx)
   If Ix = -1 Then
       Push O, Dry(J)
   Else
       O(Ix)(MgeIx) = O(Ix)(MgeIx) & Sep & Dry(J)(MgeIx)
   End If
Next
DryMge = O
End Function

Function DryMgeIx&(Dry, Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   DryMgeIx = O
   Exit Function
Nxt:
Next
DryMgeIx = -1
End Function

Function Dry_NCol%(A)
Dim O%, Dr
For Each Dr In AyNz(A)
    O = Max(O, Sz(Dr))
Next
Dry_NCol = O
End Function

Function DryPkMinus(A, B, PkIxAy&()) As Variant()
Dim AK(): AK = DrySel(A, PkIxAy)
Dim BK(): BK = DrySel(B, PkIxAy)
Dim CK(): CK = DryPkMinus(AK, BK, PkIxAy)
DryPkMinus = Dry_XWh_IxAyVy(A, PkIxAy, CK)
End Function

Function DryReOrd(Dry, PartialIxAy&()) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim Dr, O()
For Each Dr In Dry
   Push O, AyReOrd(Dr, PartialIxAy)
Next
DryReOrd = O
End Function

Function Dry_XRmv_Col(A, ColIx&) As Variant()
Dim X
For Each X In AyNz(A)
    PushI Dry_XRmv_Col, Ay_XExl_EleAt(X, ColIx)
Next
End Function

Function Dry_XRmv_ColByIxAy(A, IxAy) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), Dr
For Each Dr In A
   Push O, Ay_XExl_IxAy(Dr, IxAy)
Next
Dry_XRmv_ColByIxAy = O
End Function

Function DryRowCnt&(Dry, ColIx&, EqVal)
If Sz(Dry) = 0 Then Exit Function
Dim J&, O&, Dr
For Each Dr In Dry
   If Dr(ColIx) = EqVal Then O = O + 1
Next
DryRowCnt = O
End Function

Function Dry_Sq(A) As Variant()
Dim O(), C%, R&, Dr
Dim NC%, NR&
NC = Dry_NCol(A)
NR = Sz(A)
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
    Dr = A(R - 1)
    For C = 1 To Min(Sz(Dr), NC)
        O(R, C) = Dr(C - 1)
    Next
Next
Dry_Sq = O
End Function

Function DryStrCol(A, Optional ColIx% = 0) As String()
DryStrCol = DryColInto(A, ColIx, EmpSy)
End Function

Function DrySy(A, Optional ColIx% = 0) As String()
DrySy = DryStrCol(A, ColIx)
End Function

Function Dry_XWh_(Dry(), ColIx%, EqVal) As Variant()
Dim O()
Dim J&
For J = 0 To UB(Dry)
   If Dry(J)(ColIx) = EqVal Then Push O, Dry(J)
Next
Dry_XWh_ = O
End Function

Function Dry_XWh_CCNe(A, C1, C2) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C1) <> Dr(C2) Then PushI Dry_XWh_CCNe, Dr
Next
End Function

Sub Dry_XAss_IsEq(A(), B())
If Not DryIsEq(A, B) Then Stop
End Sub

Function Dry_WdtAy(A()) As Integer()
Dim J%
For J = 0 To Dry_NCol(A) - 1
    Push Dry_WdtAy, Ay_Wdt(DryCol(A, J))
Next
End Function

Function Dry_WdtAy1(A(), Optional MaxColWdt% = 100) As Integer()
Const CSub$ = CMod & "Dry_WdtAy"
If Sz(A) = 0 Then Exit Function
Dim O%()
   Dim Dr, UDr%, U%, V, L%, Ix&, J&
   U = -1
   For Each Dr In A
       If Not IsSy(Dr) Then
            Dim Msg$
            Msg = "This routine should call ACvFmtEachCell first so that each cell is ValCellStr as a string.  "
            Msg = Msg + "But now some Dr in given-Dry is not a StrAy"
            XThw CSub, Msg, "Dry-URow Dry-Ix-with-Err TypeName-of-Dr-of-Dry", UB(A), Ix, TypeName(Dr)
       End If
       UDr = UB(Dr)
       If UDr > U Then ReDim Preserve O(UDr): U = UDr
       If Sz(Dr) = 0 Then GoTo Nxt
       For J = 0 To UDr
           V = Dr(J)
           L = Len(V)

           If L > O(J) Then O(J) = L
       Next
Nxt:
        Ix = Ix + 1
   Next
Dim M%
    M = Limit(MaxColWdt, 1, 200)
For J = 0 To UB(O)
   If O(J) > M Then O(J) = M
Next
Dry_WdtAy1 = O
End Function

Function Dry_XWh_ColEq(A, C%, V) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C) = V Then PushI Dry_XWh_ColEq, Dr
Next
End Function

Function Dry_XWh_ColGt(A, C%, V) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    If Dr(C) > V Then PushI Dry_XWh_ColGt, Dr
Next
End Function

Function Dry_XWh_Dup(A, Optional ColIx% = 0) As Variant()
Dim Dup, Dr, O()
Dup = Ay_XWh_Dup(DryCol(A, ColIx))
For Each Dr In A
    If Ay_XHas(Dup, Dr(ColIx)) Then Push O, Dr
Next
Dry_XWh_Dup = O
End Function

Function Dry_XWh_IxAyVy(A, IxAy, Vy) As Variant()
Dim Dr
For Each Dr In A
    If Ay_IsEq(DrSel(Dr, IxAy), Vy) Then PushI Dry_XWh_IxAyVy, Dr
Next
End Function

Function DryDistCol(A() As Variant, ColIx%)
DryDistCol = Ay_XWh_Dist(DryCol(A, ColIx))
End Function

Function DryDistSy(A() As Variant, ColIx%) As String()
DryDistSy = Ay_XWh_Dist(DrySy(A, ColIx))
End Function

Function DryJnDotSy(A() As Variant) As String()
Dim Dr
For Each Dr In AyNz(A)
    PushI DryJnDotSy, JnDot(Dr)
Next
End Function
