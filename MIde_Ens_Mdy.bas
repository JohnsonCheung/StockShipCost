Attribute VB_Name = "MIde_Ens_Mdy"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_XEns_Mdy."

Function MthLin_XEns_Prv$(A$)
Lin_XAss_MthLin A, CSub
MthLin_XEns_Prv = "Private " & XRmv_Mdy(A)
End Function

Function Lin_XAss_MthLin(A, Fun$)
If Not Lin_IsMth(A) Then XThw Fun, "Given [Lin] is not MthLin", A
End Function
Function MthLin_XEns_Pub$(A$)
Lin_XAss_MthLin A, CSub
MthLin_XEns_Pub = XRmv_Mdy(A)
End Function

Sub XEns_ZPrv()
Md_XEns_ZPrv CurMd
End Sub
Sub XEns_PjZPrv()
Pj_XEns_ZPrv CurPj
End Sub

Sub Pj_XEns_ZPrv(A As VBProject)
Dim Md
For Each Md In AyNz(Pj_MdAy(A))
    Md_XEns_ZPrv CvMd(Md)
Next
End Sub

Sub XEns_ZPub()
Md_XEns_ZDashPub CurMd
End Sub

Sub Md_XEns_ZPrv(A As CodeModule)
Md_XEns_Prv_Wh A, WhMth(WhMdy:="Pub", Nm:=WhNm("^Z.+"))
End Sub

Sub Md_XEns_ZDashPub(A As CodeModule)
Md_XEns_Pub_Wh A, WhMth(WhMdy:="Prv", Nm:=WhNm("^Z_"))
End Sub

Sub MdMth_XEns_Frd(A As CodeModule, MthNm$)
Const CSub$ = CMod & "MdMth_XEns_Frd"
If Md_Ty(A) <> vbext_ct_ClassModule Then XThw CSub, "Given MdTy is not Class", "Md Ty", Md_Nm(A), Md_TyStr(A)
MdMthNm_XEns_Mdy A, MthNm, "Friend"
End Sub

Sub MdMth_XEns_Mdy(A As CodeModule, MthNm$, Mdy$)
Dim Lno
For Each Lno In AyNz(MdMthNm_LnoAy(A, MthNm))
   MdMthLno_XEns_Mdy A, Lno, Mdy
Next
End Sub

Sub MdMth_XEns_Prv(A As CodeModule, MthNm$)
MdMthNm_XEns_Mdy A, MthNm, "Private"
End Sub

Sub MdMth_XEns_Pub(A As CodeModule, MthNm$)
MdMthNm_XEns_Mdy A, MthNm
End Sub

Sub Md_XEns_Prv_Wh(A As CodeModule, B As WhMth)
Dim Ix
For Each Ix In AyNz(Src_MthIxAy(Md_Src(A), B))
    MdMthLno_XEns_Prv A, Ix + 1
Next
End Sub

Sub Md_XEns_Pub_Wh(A As CodeModule, B As WhMth)
Dim Ix
For Each Ix In Src_MthIxAy(Md_Src(A), B)
    MdMthLno_XEns_Pub A, Ix + 1
Next
End Sub
Private Function ZNewLin$(OldLin$, Mdy$)
Const CSub$ = CMod & "ZNewLin"
Dim L$
    L = XRmv_Mdy(OldLin)
    Select Case Mdy
    Case "Pub", "": ZNewLin = L
    Case "Prv":     ZNewLin = "Private " & L
    Case "Frd":     ZNewLin = "Friend " & L
    Case Else
        XThw CSub, "Given parameter [Mdy] must be ['' Pub Prv Frd]", "Mdy", Mdy
    End Select
End Function

Private Sub MdMthLno_XEns_Mdy(A As CodeModule, MthLno, Optional Mdy$)
Const CSub$ = CMod & "MdMthLno_XEns_Mdy"
Dim OldLin$
    OldLin = A.Lines(MthLno, 1)
    If Not Lin_IsMth(OldLin) Then
       XThw CSub, "Given [Md]-[MthLno]-[Lin] is not a method", "Md MthLno Lin", Md_Nm(A), MthLno, OldLin
    End If
Dim NewLin$: NewLin = ZNewLin(OldLin, Mdy)
If OldLin = NewLin Then
   Debug.Print CSub; QQ_Fmt(": Same Mdy[?] in Lin[?]", Mdy, OldLin)
   Exit Sub
End If
Md_XRpl_Lin A, MthLno, NewLin
Debug.Print CSub
Debug.Print QQ_Fmt("  Mdy[?] Of MthLno[?] of Md[?] is XEnsured", Mdy, MthLno, Md_Nm(A))
Debug.Print QQ_Fmt("  OldLin[?]", OldLin)
Debug.Print QQ_Fmt("  NewLin[?]", NewLin)
End Sub

Private Sub MdMthNm_XEns_Mdy(A As CodeModule, MthNm$, Optional Mdy$)
Dim Lno
For Each Lno In AyNz(MdMthNm_LnoAy(A, MthNm))
   MdMthLno_XEns_Mdy A, Lno, Mdy
Next
End Sub

Private Function MdMthLno_XEns_Prv(A As CodeModule, MthLno)
MdMthLno_XEns_Mdy A, MthLno, "Prv"
End Function

Private Function MdMthLno_XEns_Pub(A As CodeModule, MthLno)
MdMthLno_XEns_Mdy A, MthLno
End Function

Private Sub Z_MdMth_XEns_Mdy()
Dim M As CodeModule
Dim MthNm$
Dim Mdy$
'--
Set M = CurMd
MthNm = "Z_A"
Mdy = "Prv"
GoSub Tst
Exit Sub
Tst:
    MdMth_XEns_Mdy M, MthNm, Mdy
    Return
End Sub

Private Sub Z_MdMth_XEns_Pub()
Dim M As Mth: Set M = New_Mth(Md("ZZModule"), "YYA")
MdMth_XEns_Prv M, "ZZA": Ass Mth_MthLin(M) = "Private Property Get ZZA()"
MdMth_XEns_Pub M, "ZZA": Ass Mth_MthLin(M) = "Property Get ZZA()"
End Sub

Function Lin_IsPrvZZDash(A) As Boolean
Dim L$: L = A
If XShf_MthMdy(L) <> "Private" Then Exit Function
If XShf_MthShtTy(L) <> "Sub" Then Exit Function
If Left(L, 3) <> "ZZ_" Then Exit Function
Lin_IsPrvZZDash = True
End Function

Function Lin_IsPubZDash(A) As Boolean
Dim L$: L = A
If Not Ay_XHas(Array("", "Public"), XShf_MthMdy(L)) Then Exit Function
If XShf_MthShtTy(L) <> "Sub" Then Exit Function
Lin_IsPubZDash = Left(L, 2) = "Z_"
End Function

Function Lin_IsPubZZDash(A) As Boolean
Dim L$: L = A
If Not Ay_XHas(Array("", "Public"), XShf_MthMdy(L)) Then Exit Function
If XShf_MthShtTy(L) <> "Sub" Then Exit Function
Lin_IsPubZZDash = Left(L, 3) = "ZZ_"
End Function

Private Sub ZZ_Lin_IsPrvZZDash()
Dim L
For Each L In CurSrc
    If Lin_IsPrvZZDash(L) Then
        Debug.Print L
    End If
Next
End Sub

Private Sub ZZ_Lin_IsPubZZDash()
Dim L
For Each L In CurSrc
    If Lin_IsPubZZDash(L) Then
        Debug.Print L
    End If
Next
End Sub

Sub MdMth_XSet_Mdy(A As CodeModule, MthNm$, Mdy$)
Ass IsMdy(Mdy)
Dim I&
    I = MdMthNm_Lno(A, MthNm)
Dim L$
    L = A.Lines(I, 1)
Dim Old$
    Old = LinMdy(L)
If Mdy = Old Then Exit Sub
Dim NewL$
    Dim B$
    If Mdy <> "" Then
        B = Mdy & " "
    Else
        B = Mdy
    End If
    NewL = B & L
With A
    .DeleteLines I, 1
    .InsertLines I, NewL
End With
End Sub

Sub MdMdMthNm_XSet_Prv(A As CodeModule, MthNm$)
MdMth_XSet_Mdy A, MthNm, "Private"
End Sub

Sub MdMth_XSet_Pub(A As CodeModule, MthNm$)
MdMth_XSet_Mdy A, MthNm, ""
End Sub

Sub Mth_XSet_Mdy(A As CodeModule, MthNm$, Mdy$)
Ass IsMdy(Mdy)
Dim I&
    I = MdMthNm_Lno(A, MthNm)
Dim L$
    L = A.Lines(I, 1)
Dim Old$
    Old = LinMdy(L)
If Mdy = Old Then Exit Sub
Dim NewL$
    Dim B$
    If Mdy <> "" Then
        B = Mdy & " "
    Else
        B = Mdy
    End If
    NewL = B & L
With A
    .DeleteLines I, 1
    .InsertLines I, NewL
End With
End Sub

Sub MdMthNm_XSet_Prv(A As CodeModule, MthNm$)
Mth_XSet_Mdy A, MthNm, "Private"
End Sub

Sub MdMthNm_XSet_Pub(A As CodeModule, MthNm$)
Mth_XSet_Mdy A, MthNm, ""
End Sub

Sub Pj_XEns_ZDashMthAsPrv(A As VBProject)
Itr_XDo Pj_MdAy(A), "Md_XEns_Z3DMthAsPrivate"
End Sub

Sub Pj_XEns_ZZDashAsPrv(A As VBProject)

End Sub

Sub Pj_XEns_ZZDashAsPub(A As VBProject)
Ay_XDo Pj_MdAy(A), "Md_XEns_ZZDashAsPrv"
End Sub


Private Sub Md_XEns_ZZDashAsPrv(A As CodeModule)
Dim DNm$: DNm = Md_DNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print QQ_Fmt("Md_XEns_ZZDashAsPrv Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Lin_IsPubZZDash(L) Then
        Debug.Print L
        By = MthLin_XEns_Prv(L)
        Debug.Print QQ_Fmt("Md_XEns_ZZDashAsPrv: Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        'A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub Md_XEns_ZZDashPrvMthAsPub(A As CodeModule)
Dim DNm$: DNm = Md_DNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print QQ_Fmt("Md_XEns_ZZDashPrvMthAsPub: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Lin_IsPrvZZDash(L) Then
        By = MthLin_XEns_Pub(L)
        Debug.Print QQ_Fmt("Md_XEns_ZZDashPrvMthAsPub: Md(?) Lin(?) is change to Public: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub Z()
Z_MdMth_XEns_Mdy
Z_MdMth_XEns_Pub
MIde_XEns_Mdy:
End Sub
