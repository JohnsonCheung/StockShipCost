Attribute VB_Name = "MIde_Mth_Dup_Dta"
Option Compare Binary
Option Explicit

Function DupMthFNyGp_Dry(Ny$()) As Variant()
'Given Ny: Each Nm in Ny is FunNm:Pj_Nm.MdNm
'          It has at least 2 ele
'          Each FunNm is same
'Return: N-Dr of Fields {Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src}
'        where N = Sz(Ny)-1
'        where each-field-(*-1)-of-Dr comes from Ny(0)
'        where each-field-(*-2)-of-Dr comes from Ny(1..)

Dim Md1$, Pj1$, Nm$
    FunFNm_BrkAsg Ny(0), Nm, Pj1, Md1
Dim Mth1 As Mth
    Set Mth1 = New_Mth(Md(Pj1 & "." & Md1), Nm)
Dim Src1$
    Src1 = MthLines(Mth1)
Dim Mdy1$, Ty1$
    Mth_MthNmBrk_XAsg Mth1, Mdy1, Ty1
Dim O()
    Dim J%
    For J = 1 To UB(Ny)
        Dim Pj2$, Nm2$, Md2$
            FunFNm_BrkAsg Ny(J), Nm2, Pj2, Md2: If Nm2 <> Nm Then Stop
        Dim Mth2 As Mth
            Set Mth2 = New_Mth(Md(Pj2 & "." & Md2), Nm)
            Dim Src2$
            Src2 = MthLines(Mth2)
        Dim Mdy2$, Ty2$
            Mth_MthNmBrk_XAsg Mth2, Mdy2, Ty2

        Push O, Array(Nm, _
                    Mdy1, Ty1, Pj1, Md1, _
                    Mdy2, Ty2, Pj2, Md2, Src1, Src2, Pj1 = Pj2, Md1 = Md2, Src1 = Src2)
    Next
DupMthFNyGp_Dry = O
End Function

Function PjDupMth_Pj_Md_Mth_Dry(A As VBProject) As Variant()
Dim Dry(): Dry = PjPubMth_Pj_Md_Mth_Dry(A)
PjDupMth_Pj_Md_Mth_Dry = Dry_XWh_ColHasDup(Dry, 2)
End Function

Private Function Pj_FfnAyDupDry(A$()) As Variant()

End Function

Private Sub Z_PjPubMth_Pj_Md_Mth_Dry()
Dry_XBrw PjPubMth_Pj_Md_Mth_Dry(CurPj)
End Sub

Function PjPubMth_Pj_Md_Mth_Dry(A As VBProject) As Variant()
Dim Md As CodeModule, M, MNm$, N, Pnm$
Pnm = Pj_Nm(A)
For Each M In AyNz(PjModAy(A))
    Set Md = M
    MNm = Md_Nm(Md)
    For Each N In AyNz(Md_MthNy(Md, WhMth(WhMdy:="Pub")))
        PushI PjPubMth_Pj_Md_Mth_Dry, Array(Pnm, MNm, N)
    Next
Next
End Function

Private Sub Z_PjPubMth_Pj_Md_Mth_Kd_Mdy_Dry()
Dry_XBrw PjPubMth_Pj_Md_Mth_Kd_Mdy_Dry(CurPj)
End Sub

Function PjPubMth_Pj_Md_Mth_Kd_Mdy_Dry(A As VBProject) As Variant()
Dim Md As CodeModule, M, MNm$, Brk, Pnm$, Kd$, N$, Mdy$
Pnm = Pj_Nm(A)
For Each M In AyNz(PjModAy(A))
    Set Md = M
    MNm = Md_Nm(Md)
    For Each Brk In AyNz(Md_MthBrkAy(Md, WhMth(WhMdy:="Pub")))
        AyAsg Brk, Mdy, Kd, N
        PushI PjPubMth_Pj_Md_Mth_Kd_Mdy_Dry, Array(Pnm, MNm, N, Kd, Mdy)
    Next
Next
End Function

Function VbeDupMthDrs(A As Vbe, B As WhPjMth, Optional IsSamMthBdyOnly As Boolean, Optional IsNoSrt As Boolean) As Drs
Dim Fny$(), Dry()
Fny = SplitSsl("Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src")
Dry = VbeDupMthDry_XWh_(A, B, IsSamMthBdyOnly:=IsSamMthBdyOnly)
Set VbeDupMthDrs = New_Drs(Fny, Dry)
End Function

Function VbeDupMthDry(A As Vbe) As Variant()
'Dim B(): B = Vbe_MthDry(A)
'Dim Ny$(): Ny = DryStrCol(B, 2)
'Dim N1$(): N1 = Ay_XWh_Dup(Ny)
'    N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
'Dim GpAy()
'    GpAy = DupMthFNy_GpAy(N1)
'    If Sz(GpAy) = 0 Then Exit Function
'Dim O()
'    Dim Gp
'    For Each Gp In GpAy
'        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
'    Next
'VbeDupMthDry = O
End Function

Function VbeDupMthDry_XWh_(A As Vbe, B As WhPjMth, Optional IsSamMthBdyOnly As Boolean) As Variant()
'Dim N$(): 'N = VbeFunFNm(A)
'Dim N1$(): ' N1 = MthNyWhDup(N)
'    If IsSamMthBdyOnly Then
'        N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
'    End If
'Dim GpAy()
'    GpAy = DupMthFNy_GpAy(N1)
'    If Sz(GpAy) = 0 Then Exit Function
'Dim O()
'    Dim Gp
'    For Each Gp In GpAy
'        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
'    Next
'VbeDupMthDry_XWh_ = O
End Function

Private Function VbePubMth_Pj_Md_Mth_Dry(A As VBProject) As Variant()
Dim Pj As VBProject, Dry(), P
For Each P In Vbe_PjAy(A)
    Set Pj = P
    Dry = Dry_XIns_Col(PjPubMth_Pj_Md_Mth_Dry(Pj), Pj_Nm(Pj))
    PushIAy VbePubMth_Pj_Md_Mth_Dry, Dry
Next
End Function


Private Sub Z_PjDupMth_Pj_Md_Mth_Dry()
Brw Dry_Fmtss(Dry_XSrt(PjDupMth_Pj_Md_Mth_Dry(CurPj), 2))
Brw Dry_Fmtss(PjPubMth_Pj_Md_Mth_Dry(CurPj))
End Sub


Private Sub Z()
Z_PjDupMth_Pj_Md_Mth_Dry
Z_PjPubMth_Pj_Md_Mth_Dry
Z_PjPubMth_Pj_Md_Mth_Kd_Mdy_Dry
MIde_Mth_Dup_Dta:
End Sub
