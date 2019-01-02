Attribute VB_Name = "MIde_Ens_SubZ"
Option Compare Binary
Option Explicit

Sub XEns_SubZZZ_PJ()
Pj_XEns_SubZZZ CurPj
Pj_XEns_ZPrv CurPj
End Sub

Sub XEns_SubZZZ_MD()
Md_XEns_SubZZZ CurMd 'XEnsure Sub Z()
Md_XEns_ZPrv CurMd 'XEnsure all Z_XXX() as Private
End Sub

Private Function MthNy_Lines_SUBZ$(MthNy$()) ' Sub Z() bodylines
Dim O$()
PushI O, "Private Sub Z()"
PushIAy O, Ay_XSrt(MthNy)
PushI O, "End Sub"
MthNy_Lines_SUBZ = JnCrLf(O)
End Function

Private Function Md_SubZZZ_ACT$(A As CodeModule)
Md_SubZZZ_ACT = MdMthNm_Lines(A, "Z") & vbCrLf & vbCrLf & MdMthNm_Lines(A, "ZZ")
End Function

Private Function Md_SubZ_EPT$(A As CodeModule)
'SubZ is [Mth-`Sub Z()`-Lines], each line is calling a Z_XXX, where Z_XXX is a testing function
Dim ZMthNy$()
    ZMthNy = Ay_XExl_Patn(Md_MthNy(A, WhMth(WhMdy:="Prv", WhKd:="Sub", Nm:=WhNm("^Z_"))), "__")
Md_SubZ_EPT = MthNy_Lines_SUBZ(ZMthNy)
End Function
Private Function Md_SubZZZ_EPT$(A As CodeModule)
Md_SubZZZ_EPT = Md_SubZ_EPT(A) & vbCrLf & vbCrLf & Md_SubZZ_EPT(A)
End Function

Private Sub Md_XEns_SubZZZ(A As CodeModule)
Ept = Md_SubZZZ_EPT(A)
Act = Md_SubZZZ_ACT(A)
If Act = Ept Then Exit Sub
'Brw Ept
'Stop
'Lines_Cmp Act, Ept, "Act-SubZ & SubZZ", "Ept"
MdMthNm_XRmv A, "Z"
MdMthNm_XRmv A, "ZZ"
If Ept <> "" Then
    Md_XApp_Lines A, vbCrLf & Ept
End If
End Sub

Private Sub Pj_XEns_SubZZZ(A As VBProject)
Dim I
For Each I In Pj_MdAy(A)
    Debug.Print Md_Nm(CvMd(I))
    Md_XEns_SubZZZ CvMd(I)
Next
End Sub

Private Sub Z_Md_SubZZ_EPT()
Dim A As CodeModule
GoSub Cas2
GoSub Cas1
Exit Sub
Cas2:
    Set A = Md("MVb_Dic")
    'ConstFunNm_XUpd_Src_BY_STR "Z_Md_SubZZ__Ept2", Md_SubZZ(A): Return
    Ept = Z_Md_SubZZ__Ept2
    GoSub Tst
    Return
Cas1:
    Set A = CurMd
    ConstFunNm_XUpd_Src_BY_STR "Z_Md_SubZZ__Ept1", Md_SubZZ_EPT(A): Return
    Ept = Z_Md_SubZZ__Ept1
    GoSub Tst
    Return
Tst:
    Act = Md_SubZZ_EPT(A)
    C
    Return
End Sub


Private Property Get Z_Md_SubZZ__Ept2$()
Const A_1$ = "Private Sub ZZ()" & _
vbCrLf & "Dim A As Variant" & _
vbCrLf & "Dim B As Dictionary" & _
vbCrLf & "Dim C() As Dictionary" & _
vbCrLf & "Dim D$" & _
vbCrLf & "Dim E$()" & _
vbCrLf & "Dim F As Boolean" & _
vbCrLf & "Dim G()" & _
vbCrLf & "CvDic A" & _
vbCrLf & "CvDicAy A" & _
vbCrLf & "DicAddAy B, C" & _
vbCrLf & "DicAddKeyPfx B, A" & _
vbCrLf & "Dic_XIup B, D, A, D" & _
vbCrLf & "DicAllKeyIsNm B" & _
vbCrLf & "DicAllKeyIsStr B" & _
vbCrLf & "DicAllValIsStr B" & _
vbCrLf & "DicAyKy C" & _
vbCrLf & "DicByDry A" & _
vbCrLf & "DicClone B" & _
vbCrLf & "DicDr B, E"

Const A_2$ = "DicDRs_Fny F" & _
vbCrLf & "DicIntersect B, B" & _
vbCrLf & "DicIsEmp B" & _
vbCrLf & "Dic_IsEq B, B" & _
vbCrLf & "Dic_IsEq_XAss B, B, D, D, D" & _
vbCrLf & "DicIsLinesDic B" & _
vbCrLf & "DicIsStrDic B" & _
vbCrLf & "DicKeySy B" & _
vbCrLf & "DicKVLy B" & _
vbCrLf & "DicKyJnVal B, A, D" & _
vbCrLf & "DicKySy B, E" & _
vbCrLf & "DicLblLy B, D" & _
vbCrLf & "DicLines B" & _
vbCrLf & "DicLy1 B" & _
vbCrLf & "DicLy2 B" & _
vbCrLf & "DicLy2__1 D, D" & _
vbCrLf & "DicMap B, D" & _
vbCrLf & "DicMaxValSz B" & _
vbCrLf & "DicMge B, D, G" & _
vbCrLf & "Dic_XMinus B, B"

Const A_3$ = "DicSelIntoAy B, E" & _
vbCrLf & "DicSelIntoSy B, E" & _
vbCrLf & "DicStrKy B" & _
vbCrLf & "DicSwapKV B" & _
vbCrLf & "DicTy B" & _
vbCrLf & "DicTyBrw B" & _
vbCrLf & "DicVal B, A" & _
vbCrLf & "LikssDicKey B, A" & _
vbCrLf & "MapStrDic A" & _
vbCrLf & "MayDicVal B, A" & _
vbCrLf & "End Sub"

Z_Md_SubZZ__Ept2 = A_1 & vbCrLf & A_2 & vbCrLf & A_3
End Property

Private Sub Z_Md_SubZZZ_EPT()
Dim A As CodeModule
'
Ept = Z_Md_SubZZZ_EPT__Ept1
Set A = CurMd
GoSub Tst
Exit Sub
Tst:
    Act = Md_SubZZZ_EPT(A)
    C
    Return
End Sub

Private Property Get Z_Md_SubZZZ_EPT__Ept1$()
Const A_1$ = "Private Sub Z()" & _
vbCrLf & "Z_SubZZLines" & _
vbCrLf & "Z_Md_SubZZZ_EPT" & _
vbCrLf & "End Sub" & _
vbCrLf & "" & _
vbCrLf & "Private Sub ZZ()" & _
vbCrLf & "A5" & _
vbCrLf & "XEns_SubZZZ_PJ" & _
vbCrLf & "XEns_SubZZZ_MD" & _
vbCrLf & "End Sub"

Z_Md_SubZZZ_EPT__Ept1 = A_1
End Property

Private Sub ZZ()
XEns_SubZZZ_PJ
XEns_SubZZZ_MD
End Sub

Private Sub Z()
Z_Md_SubZZ_EPT
Z_Md_SubZZZ_EPT
End Sub

Private Property Get Z_Md_SubZZ__Ept1$()
Const A_1$ = "Private Sub ZZ()" & _
vbCrLf & "A5" & _
vbCrLf & "XEns_SubZZZ_PJ" & _
vbCrLf & "XEns_SubZZZ_MD" & _
vbCrLf & "End Sub"

Z_Md_SubZZ__Ept1 = A_1
End Property

