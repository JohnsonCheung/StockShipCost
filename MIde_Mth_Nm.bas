Attribute VB_Name = "MIde_Mth_Nm"
Option Compare Binary
Option Explicit

Property Get CurMd_MthNy() As String()
CurMd_MthNy = Md_MthNy(CurMd)
End Property

Property Get CurMth_DNm$()
Dim M$: M = CurMthNm
If M = "" Then Exit Property
CurMth_DNm = CurMd_Nm & "." & M
End Property

Function CurPj_MthNy(Optional A As WhPjMth, Optional AddPjNm As Boolean) As String()
CurPj_MthNy = Pj_MthNy(CurPj, A, AddPjNm)
End Function

Function MthNy(Optional A As WhPjMth) As String()
MthNy = CurVbe_MthNy(A)
End Function

Function MthNy_Pub(Optional A As WhPjMth) As String()
MthNy_Pub = CurVbe_MthNy_Pub(A)
End Function

Function CurVbe_MthNy(Optional A As WhPjMth) As String()
CurVbe_MthNy = Vbe_MthNy(CurVbe, A)
End Function

Function CurVbe_MthNy_Pub(Optional A As WhPjMth) As String()
CurVbe_MthNy_Pub = Vbe_MthNy_Pub(CurVbe, A)
End Function

Function DDNmMth(MthDDNm$) As Mth
Dim M As CodeModule
Dim Nm$
Dim Ny$(): Ny = Split(MthDDNm, ".")
Select Case Sz(Ny)
Case 1: Nm = Ny(0): Set M = CurMd
Case 2: Nm = Ny(1): Set M = Md(Ny(0))
Case 3: Nm = Ny(2): Set M = Pj_Md(Pj(Ny(0)), Ny(1))
Case Else: Stop
End Select
Set DDNmMth = New_Mth(M, Nm)
End Function

Function Fb_MthNy(A) As String()
Fb_MthNy = Vbe_MthNy(Fb_Acs(A).Vbe)
End Function

Function Lin_MthNm$(A)
Lin_MthNm = MthNmBrk_Nm(Lin_MthNmBrk(A))
End Function

Function Lin_PrpNm$(A)
Dim L$
L = XRmv_Mdy(A)
If XShf_Kd(L) <> "Property" Then Exit Function
Lin_PrpNm = XTak_Nm(L)
End Function

Function MdMthNm_Dft$(Optional A As CodeModule, Optional MthNm$)
If MthNm = "" Then
   MdMthNm_Dft = Md_CurMthNm(DftMd(A))
Else
   MdMthNm_Dft = A
End If
End Function

Function Md_MthDDNy(A As CodeModule, Optional AddMdNm As Boolean) As String()
Dim O$()
O = Src_MthDDNy(Md_Src(A))
If AddMdNm Then O = Ay_XAdd_Pfx(O, Md_Nm(A) & ".")
Md_MthDDNy = O
End Function

Function Md_MthNy(A As CodeModule, Optional B As WhMth) As String()
Md_MthNy = Src_MthNy(Md_BdyLy(A), B)
End Function

Function MthDDNy_XWh(A$(), B As WhMth) As String()
If IsNothing(B) Then
    MthDDNy_XWh = A
    Exit Function
End If
Dim N
For Each N In AyNz(A)
    If MthDDNm_IsSel(N, B) Then
        PushI MthDDNy_XWh, N
    End If
Next
End Function

Function Mth_MthDNm_Nm$(A)
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$
Select Case Sz(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
Mth_MthDNm_Nm = Nm
End Function

Function MthDotNTM$(MthDot$)
'MthDot is a string with last 3 seg as Mdy.Ty.Nm
'MthNTM is a string with last 3 seg as Nm:Ty.Mdy
Dim Ay$(), Nm$, Ty$, Mdy$
Ay = SplitDot(MthDot)
'AyAsg AyPop(Ay), Ay, Nm
'AyAsg AyPop(Ay), Ay, Ty
'AyAsg AyPop(Ay), Ay, Mdy
Push Ay, QQ_Fmt("?:?.?", Nm, Ty, Mdy)
MthDotNTM = JnDot(Ay)
End Function

Function MthNm$(A As Mth)
MthNm = A.Nm
End Function

Function MthNmMd(A$) As CodeModule '
Dim O As CodeModule
Set O = CurMd
If Md_XHas_MthNm(O, A) Then Set MthNmMd = O: Exit Function
Dim N$
N = MthFul(A)
If N = "" Then
    Debug.Print QQ_Fmt("Mth[?] not found in any Pj")
    Exit Function
End If
Set MthNmMd = Md(N)
End Function

Function Src_MthNy(A$(), Optional B As WhMth) As String()
Src_MthNy = Ay_XWh_Dist(Ay_XTak_BefDot(MthDDNy_XWh(Src_MthDDNy(A), B)))
End Function


Private Sub Z_Fb_MthNy()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
'    For Each Fb In AppFbAy
        PushAy O, Fb_MthNy(Fb)
'    Next
    Brw O
    Return
X_BrwOne:
'    Brw Fb_MthNy(AppFbAy()(0))
    Return
End Sub

Private Sub Z_Lin_MthNm()
GoTo ZZ
Dim A$
A = "Function Lin_MthNm$(A)": Ept = "Lin_MthNm.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = Lin_MthNm(A)
    C
    Return
ZZ:
    Dim O$(), L, P, M
    For Each P In Vbe_PjAy(CurVbe)
        For Each M In Pj_MdAy(CvPj(P))
            For Each L In Md_BdyLy(CvMd(M))
                PushNonBlankStr O, Lin_MthNm(CStr(L))
            Next
        Next
    Next
    Brw O
End Sub

Private Sub Z_MdMthNm_Dft()
Dim I, Md As CodeModule
For Each I In Pj_MdAy(CurPj)
   Md_XShw CvMd(I)
   Debug.Print Md_Nm(Md), MdMthNm_Dft(Md)
Next
End Sub

Private Sub Z_Src_MthNy()
Brw Src_MthNy(CurSrc)
End Sub

Function Md_MthNy_ZZDASH(A As CodeModule) As String()
Dim J%, O$(), L$, L1$, Is_ZFun As Boolean
For J = 1 To A.CountOfLines
    Is_ZFun = True
    L = A.Lines(J, 1)
    Select Case True
    Case XHas_Pfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = XRmv_Pfx(L, "Sub ")
    Case XHas_Pfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = XRmv_Pfx(L, "Sub ")
    Case Else:
        Is_ZFun = False
    End Select

    If Is_ZFun Then
        Push O, XTak_Nm(L1)
    End If
Next
Md_MthNy_ZZDASH = O
End Function

Function Md_MthNy_Pub(A As CodeModule) As String()
If Md_Ty(A) <> vbext_ct_StdModule Then XThw CSub, "Given Md should be Std", "Md MdTy", Md_Nm(A), Md_TyStr(A)
Md_MthNy_Pub = Ay_XWh_Dist(Src_MthNy(Md_Src(A), WhMth("Pub")))
End Function

Private Sub Z()
Z_Fb_MthNy
Z_Lin_MthNm
Z_MdMthNm_Dft
Z_Src_MthNy
MIde_Mth_Nm:
End Sub

Function Pj_MthNy(A As VBProject, Optional B As WhMdMth, Optional AddPjNm As Boolean) As String()
Dim Md As CodeModule, I, N$, Ny$()
If AddPjNm Then N = A.Name & "."
Dim M As WhMth
If Not IsNothing(B) Then Set M = B.Mth
For Each I In AyNz(Pj_MdAy(A, WhMdMth_WhMd(B)))
    Set Md = I
    Ny = MthDDNy_XWh(Md_MthDDNy(Md), M)
    Ny = Ay_XAdd_Pfx(Ny, N & Md_Nm(Md) & ".")
    PushAyNoDup Pj_MthNy, Ny
Next
End Function


