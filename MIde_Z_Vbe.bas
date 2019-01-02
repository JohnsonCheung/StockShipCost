Attribute VB_Name = "MIde_Z_Vbe"
Option Compare Binary
Option Explicit

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function

Property Get PjNy() As String()
PjNy = Vbe_PjNy(CurVbe)
End Property

Sub Vbe_XDmp_PjIsSaved(A As Vbe)
Dim I As VBProject
For Each I In A.VBProjects
    Debug.Print I.Saved, I.BuildFileName
Next
End Sub

Function Vbe_Pj(A As Vbe, PjNm$) As VBProject
Set Vbe_Pj = A.VBProjects(PjNm)
End Function

Function VbePjFfn_Pj(A As Vbe, Ffn) As VBProject
Dim I
For Each I In A.VBProjects ' Cannot use Vbe_PjAy(A), should use A.VBProjects
                           ' due to Vbe_PjAy(X).FileName gives error
                           ' but (Pj in A.VBProjects).FileName is OK
    Debug.Print Pj_Ffn(CvPj(I))
    If StrComp(Pj_Ffn(CvPj(I)), Ffn) Then
        Set VbePjFfn_Pj = I
        Exit Function
    End If
Next
End Function

Function VbePj_Ffn(A As Vbe, Ffn) As VBProject
Dim I As VBProject
For Each I In A.VBProjects
    If Pj_Ffn(I) = Ffn Then Set VbePj_Ffn = I: Exit Function
Next
End Function

Function VbePj_FfnPj(A As Vbe, Ffn) As VBProject
Dim P As VBProject
For Each P In A.VBProjects
    If StrComp(Pj_Ffn(P), Ffn, vbTextCompare) = 0 Then
        Set VbePj_FfnPj = P
        Exit Function
    End If
Next
End Function

Function VbePj_MdDry(A As Vbe) As Variant()
Dim O(), P, C, Pnm$, Pj As VBProject
For Each P In Vbe_PjAy(A)
    Set Pj = P
    Pnm = Pj_Nm(Pj)
    For Each C In PjCmpAy(Pj)
        Push O, Array(Pnm, CvCmp(C).Name)
    Next
Next
VbePj_MdDry = O
End Function

Function VbePj_MdFmt(A As Vbe) As String()
VbePj_MdFmt = Dry_Fmtss(VbePj_MdDry(A))
End Function

Sub VbePj_MdFmtBrw(A As Vbe)
Brw VbePj_MdFmt(A)
End Sub

Sub VbeSav(A As Vbe)
Itr_XDo A.VBProjects, "Pj_XSav"
End Sub

Function VbeVisWinCnt%(A As Vbe)
VbeVisWinCnt = Itr_Cnt_PrpTrue(A.Windows, "Visible")
End Function

Sub Vbe_Compile(A As Vbe)
Itr_XDo A.VBProjects, "Pj_XCompile"
End Sub

Function Vbe_MthKy(A As Vbe, Optional IsWrap As Boolean) As String()
Dim O$(), I
For Each I In Vbe_PjAy(A)
    PushAy O, Pj_MthKy(CvPj(I), IsWrap)
Next
Vbe_MthKy = O
End Function

Function Vbe_MthLinDry(A As Vbe, Optional B As WhPjMth) As Variant()
Dim P, WhMdMth As WhMdMth, WhPj As WhNm
Set WhMdMth = WhPjMth_WhMdMth(B)
Set WhPj = WhPjMth_WhPj(B)
For Each P In AyNz(Vbe_PjAy(A, WhPj))
    PushAy Vbe_MthLinDry, Pj_MthLinDry(CvPj(P), WhMdMth)
Next
End Function

Function Vbe_MthLinDryWP(A As Vbe) As Variant()
Dim P
For Each P In AyNz(Vbe_PjAy(A))
    PushIAy Vbe_MthLinDryWP, Pj_MthLinDry_WrapPm(CvPj(P))
Next
End Function

Function Vbe_MthNy(A As Vbe, Optional B As WhPjMth) As String()
Dim I, W As WhMdMth
Set W = WhPjMth_MdMth(B)
For Each I In AyNz(Vbe_PjAy(A, WhPjMth_WhNm(B)))
    PushIAy Vbe_MthNy, Pj_MthNy(CvPj(I), W)
Next
End Function

Function Vbe_MthNy_Pub(A As Vbe, Optional B As WhPjMth) As String()
Vbe_MthNy_Pub = Vbe_MthNy(A, WhPjMth_XAdd_Pub(B))
End Function

Function Vbe_Mth_MdDNm$(A As Vbe, MthNm$)
Dim Pj As VBProject, P, M, Md As CodeModule
For Each P In AyNz(Vbe_PjAy(A))
    Set Pj = P
    For Each M In Pj_MdAy(Pj)
        Set Md = M
        If Md_XHas_MthNm(Md, MthNm) Then Vbe_Mth_MdDNm = Md_DNm(Md) & "." & MthNm: Exit Function
    Next
Next
End Function

Function Vbe_PjAy(A As Vbe, Optional B As WhNm) As VBProject()
Vbe_PjAy = Itr_XWh_NmInto(A.VBProjects, B, Vbe_PjAy)
End Function

Function Vbe_PjFfnAy(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNonBlankStr Vbe_PjFfnAy, Pj_Ffn(P)
Next
End Function

Function Vbe_PjNy(A As Vbe, Optional B As WhNm) As String()
Vbe_PjNy = Itr_Ny(Vbe_PjAy(A, B))
End Function

Function Vbe_PjNyWh(A As Vbe, B As WhNm) As String()
Vbe_PjNyWh = Ay_XWh_Nm(Vbe_PjNy(A), B)
End Function

Function Vbe_PjNyWhMd(A As Vbe, MdPatn$) As String()
Dim I, Re As New RegExp
Re.Pattern = MdPatn
For Each I In Vbe_PjAy(A)
    If Pj_XHas_CmpNm_WhRe(CvPj(I), Re) Then
        Push Vbe_PjNyWhMd, CvPj(I).Name
    End If
Next
End Function

Function Vbe_QPj_FST(A As Vbe) As VBProject
Dim I
For Each I In A.VBProjects
    If XTak_FstChr(CvPj(I).Name) = "Q" Then
        Set Vbe_QPj_FST = I
        Exit Function
    End If
Next
End Function

Function Vbe_Src(A As Vbe) As String()
Dim P
For Each P In A.VBProjects
    PushAy Vbe_Src, Pj_Src(CvPj(P))
Next
End Function

Function Vbe_SrcPth(A As Vbe)
Dim Pj As VBProject:
Set Pj = Vbe_QPj_FST(A)
Dim Ffn$: Ffn = Pj_Ffn(Pj)
If Ffn = "" Then Exit Function
Vbe_SrcPth = Ffn_Pth(Pj.FileName)
End Function

Function Vbe_MthWb(A As Vbe) As Workbook
Set Vbe_MthWb = Wb_XVis(Ws_Wb(Vbe_MthWs(A)))
End Function

Sub Vbe_XBrw_SrcPth(A As Vbe)
Pth_XBrw Vbe_SrcPth(A)
End Sub

Sub Vbe_XBrw_SrtRptLy(A As Vbe)
Brw Vbe_SrtRptLy(A)
End Sub

Sub Vbe_XExp(A As Vbe)
Dim P
For Each P In Vbe_PjAy(A)
    Pj_XExp CvPj(P)
Next
End Sub

Function Vbe_XHas_Bar(A As Vbe, Nm$) As Boolean
Vbe_XHas_Bar = Itr_XHas_Nm(A.CommandBars, Nm)
End Function

Function Vbe_XHas_Pj(A As Vbe, Pj_Nm) As Boolean
Vbe_XHas_Pj = Itr_XHas_Nm(A.VBProjects, Pj_Nm)
End Function

Function Vbe_XHas_Pj_Ffn(A As Vbe, Ffn) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If Pj_Ffn(P) = Ffn Then Vbe_XHas_Pj_Ffn = True: Exit Function
Next
End Function

Sub Vbe_XSrt(A As Vbe)
Dim I
For Each I In Vbe_PjAy(A)
    Pj_XSrt CvPj(I)
Next
End Sub

Function Vbe_SrtRptLy(A As Vbe) As String()
Dim Ay() As VBProject: Ay = Vbe_PjAy(A)
Dim O$(), I, M As VBProject
For Each I In Ay
    Set M = I
    PushIAy O, Pj_SrtRptLy(M)
Next
Vbe_SrtRptLy = O
End Function

Private Sub ZZ_Vbe_XDmp_PjIsSaved()
Vbe_XDmp_PjIsSaved CurVbe
End Sub

Private Sub ZZ_VbeFunPfx()
'D Vbe_MthPfx(CurVbe)
End Sub

Private Sub ZZ_Vbe_MthNy()
Brw Vbe_MthNy(CurVbe)
End Sub

Private Sub ZZ_Vbe_MthNyWh()
Brw Vbe_MthNy(CurVbe)
End Sub

Private Sub ZZ()
Dim A
Dim B As Vbe
Dim C$
Dim D As Boolean
Dim E As WhPjMth
Dim F As WhNm
Dim XX
CvVbe A
VbePjFfn_Pj B, A
VbePj_Ffn B, A
VbePj_FfnPj B, A
VbePj_MdDry B
VbePj_MdFmt B
VbePj_MdFmtBrw B
VbeSav B
VbeVisWinCnt B
Vbe_Compile B
Vbe_MthKy B, D
Vbe_MthNy B, E
Vbe_MthNy_Pub B, E
Vbe_Mth_MdDNm B, C
Vbe_PjAy B, F
Vbe_PjFfnAy B
Vbe_PjNy B, F
Vbe_PjNyWh B, F
Vbe_PjNyWhMd B, C
Vbe_QPj_FST B
Vbe_Src B
Vbe_SrcPth B
Vbe_MthWb B
Vbe_XBrw_SrcPth B
Vbe_XExp B
Vbe_XHas_Bar B, C
Vbe_XHas_Pj B, A
Vbe_XHas_Pj_Ffn B, A
Vbe_XSrt B
XX = PjNy()
End Sub

Private Sub Z()
End Sub
