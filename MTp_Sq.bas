Attribute VB_Name = "MTp_Sq"
Option Compare Binary
Option Explicit
Public Const SqTpBlkTyss$ = "ER PM SW SQ RM"
Public Const Msg_Sq_1_NotInEDic = "These items not found in ExprDic [?]"
Public Const Msg_Sq_1_MustBe1or0 = "For %?xxx, 2nd term must be 1 or 0"

Private Sub AAMain()
Z_SqTp_SqyRslt
End Sub

Private Function BlkAy_LnxAy(A() As Blk, BlkTyStr$) As Lnx()
Dim J%
For J = 0 To UB(A)
    If A(J).BlkTyStr = BlkTyStr Then BlkAy_LnxAy = A(J).Gp.LnxAy: Exit Function
Next
End Function

Function BlkAy_XWh_TySelGp(A() As Blk, BlkTyStr$) As Gp()
Dim J%
For J = 0 To UB(A)
    With A(J)
        If .BlkTyStr = BlkTyStr Then
            PushObj BlkAy_XWh_TySelGp, A(J).Gp
        End If
    End With
Next
End Function

Private Function GpAy_BlkAy1(A() As Gp) As Blk()
Dim I
For Each I In AyNz(A)
    PushObj GpAy_BlkAy1, Gp_Blk(CvGp(I))
Next
End Function

Function GpAy_XRmv_Rmk(A() As Gp) As Gp()
'Dim J%, O() As Gp, M As Gp
'For J = 0 To UB(A)
'    M = Gp_XRmv_Rmk(A(J))
'    If Sz(M.LnxAy) > 0 Then
'        PushObj O, M
'    End If
'Next
'GpAy_XRmv_Rmk = O
End Function

Function Gp_Blk(A As Gp) As Blk
Set Gp_Blk = New Blk
With Gp_Blk
    .BlkTyStr = Gp_BlkTyStr(A)
    Set .Gp = A
End With
End Function

Function Gp_BlkTyStr$(A As Gp)
Dim Ly$(): Ly = Gp_Ly(A)
Dim O$
Select Case True
Case Ly_IsPm(Ly): O = "PM"
Case Ly_IsSw(Ly): O = "SW"
Case Ly_IsRm(Ly): O = "RM"
Case Ly_IsSq(Ly): O = "SQ"
Case Else: O = "ER"
End Select
Gp_BlkTyStr = O
End Function

Function LnxAy_BlkTyStr$(A() As Lnx)
Dim Ly$(): ' Ly = LnxAy_ZIs(Ly)
Dim O$
Select Case True
Case Ly_IsPm(Ly): O = "PM"
Case Ly_IsSw(Ly): O = "SW"
Case Ly_IsRm(Ly): O = "RM"
Case Ly_IsSq(Ly): O = "SQ"
Case Else: O = "ER"
End Select
LnxAy_BlkTyStr = O
End Function

Function LnxAy_LnxAyRslt_DUP_KEY(A() As Lnx) As LnxAyRslt

End Function

Function LnxAy_LnxAyRslt_PERCENTAGE_PFX(A() As Lnx) As LnxAyRslt

End Function

Function Ly_BlkTyStr$(A$())
Dim O$
Select Case True
Case Ly_IsPm(A): O = "PM"
Case Ly_IsSw(A): O = "SW"
Case Ly_IsRm(A): O = "RM"
Case Ly_IsSq(A): O = "SQ"
Case Else: O = "ER"
End Select
Ly_BlkTyStr = O
End Function

Function Ly_GpAy(Ly$()) As Gp()
Dim O() As Gp, J&, LnxAy() As Lnx, M As Lnx
For J = 0 To UB(Ly)
    Dim Lin$
    Lin = Ly(J)
    If XHas_Pfx(Lin, "==") Then
        If Sz(LnxAy) > 0 Then
            PushObj O, New_Gp(LnxAy)
        End If
        Erase LnxAy
    Else
        PushObj LnxAy, Lnx(J, Lin)
    End If
Next
If Sz(LnxAy) > 0 Then
    PushObj O, New_Gp(LnxAy)
End If
Ly_GpAy = O
End Function

Private Function Ly_IsPm(A$()) As Boolean
Ly_IsPm = LyHasMajPfx(A, "%")
End Function

Private Function Ly_IsRm(A$()) As Boolean
Ly_IsRm = Sz(A) = 0
End Function

Private Function Ly_IsSq(A$()) As Boolean
If Sz(A) <> 0 Then Exit Function
Dim L$: L = A(0)
Dim Sy$(): Sy = Ssl_Sy("?SEL SEL ?SELDIS SELDIS UPD DRP")
If XHas_PfxAy(L, Sy) Then Ly_IsSq = True: Exit Function
End Function

Private Function Ly_IsSw(A$()) As Boolean
Ly_IsSw = LyHasMajPfx(A, "?")
End Function

Private Property Get Rslt_1() As String()
'Return a split-of-SwLnxAy-and-ErLy as SwLnxAyErLy
'by if B_Ay(..).ErLy has Er
'       then put into ErLy    (E$())
'       else put into SwLnxAy (O() As SwLnx)
'Dim E$(), O() As SwLnx
'Dim J%, Er$()
'For J = 0 To U
'    Er = B_Ay(J).ErLy
'    If Ay_IsEmp(Er) Then
'        PushObj O, B_Ay(J)
'    Else
'        PushAy E, Er
'    End If
'Next
'With Rslt_1
'    .ErLy = E
'    .SwLnxAy = O
'End With
End Property

Private Function SqTp_BlkAy(SqTp$) As Blk()
Dim Ly$():            Ly = SplitCrLf(SqTp)
Dim G() As Gp:         G = Ly_GpAy(Ly)
Dim G1() As Gp:       G1 = GpAy_XRmv_Rmk(G)
SqTp_BlkAy = GpAy_BlkAy1(G1)
End Function

Function SqTp_SqyRslt(SqTp$) As SqyRslt
Dim B() As Blk: B = SqTp_BlkAy(SqTp)
Dim Pm As Dictionary
Dim SwR As SwRslt
Dim SqR As SqyRslt
Dim PmR As PmRslt
PmR = PmBlkAy_PmRslt(BlkAy_LnxAy(B, "PM"))
Set Pm = PmR.Pm
SwR = LnxAy_SwRslt(BlkAy_LnxAy(B, "SW"), Pm)
SqR = SqBlkAy_SqyRslt(BlkAy_XWh_TySelGp(B, "SQ"), Pm, SwR.StmtSw, SwR.FldSw)
Dim O As SqyRslt
O.Sqy = SqR.Sqy
O.Er = AyAp_XAdd( _
    XChk_ErBlk(B), _
    XChk_ExcessSwBlk(B), _
    XChk_ExcessPmBlk(B), _
    PmR.Er, SwR.Er, SqR.Er)
End Function

Private Function WIsSw(Ly$()) As Boolean
WIsSw = LyHasMajPfx(Ly, "?")
End Function

Function XChk_ErBlk(A() As Blk) As String()

End Function

Private Function XChk_ExcessPmBlk(A() As Blk) As String()

End Function

Private Function XChk_ExcessSwBlk(A() As Blk) As String()
Dim I, B() As Blk
For Each I In AyNz(A)
    If CvBlk(I).BlkTyStr = "SW" Then
    End If
    
    
Next
End Function

Private Property Get ZZSqTp$()
Static X$
'If X = "" Then X = Md_ResStr(Md("W01SqTp"), "SqTp")
ZZSqTp = X
End Property

Private Property Get ZZSqTpLy() As String()
ZZSqTpLy = SplitCrLf(ZZSqTp)
End Property

Private Sub Z_SqTp_SqyRslt()
Dim SqTp$, Act As SqyRslt, Ept As SqyRslt
'--
SqTp = Samp_SqTp
GoSub Tst
Exit Sub
Tst:
    Act = SqTp_SqyRslt(SqTp)
    C
    Ass SqyRslt_IsEq(Act, Ept)
    Return
End Sub
Function SqyRslt_IsEq(A As SqyRslt, B As SqyRslt) As Boolean
Stop '
End Function
Private Sub ZZ()
Dim A() As Blk
Dim B$
Dim C As Gp
Dim D() As Gp
Dim E() As Lnx
Dim F%()
Dim G$()
Dim XX
BlkAy_XWh_TySelGp A, B
GpAy_XRmv_Rmk D
Gp_Blk C
Gp_BlkTyStr C
LnxAy_BlkTyStr E
LnxAy_LnxAyRslt_DUP_KEY E
LnxAy_LnxAyRslt_PERCENTAGE_PFX E
Ly_GpAy G

XChk_ErBlk A
End Sub

Private Sub Z()
Z_SqTp_SqyRslt
End Sub
