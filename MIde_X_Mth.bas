Attribute VB_Name = "MIde_X_Mth"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_X_Mth."
Function New_Mth(A As CodeModule, MthNm) As Mth
Set New_Mth = New Mth
With New_Mth
    Set .Md = A
    .Nm = MthNm
End With
End Function

Sub MthAy_XMov(A() As Mth, ToMd As CodeModule)
Ay_XDoXP A, "Mth_XMov", ToMd
End Sub

Function Mth_MthDNm$(A As Mth)
Mth_MthDNm = Md_DNm(A.Md) & "." & A.Nm
End Function
Function MthFul$(MthNm$)
MthFul = Vbe_Mth_MdDNm(CurVbe, MthNm)
End Function

Property Get MthKeyDrFny() As String()
MthKeyDrFny = Ssl_Sy("Pj_Nm MdNm Priority Nm Ty Mdy")
End Property

Function MthKy_Sq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
Sq_Set_Row O, 1, MthKeyDrFny
For J = 0 To UB(A)
    Sq_Set_Row O, J + 2, Split(A(J), ":")
Next
MthKy_Sq = O
End Function


Function CvMth(A) As Mth
Set CvMth = A
End Function

Function Mth_Lno&(A As Mth)
Mth_Lno = MdMthNm_Lno(A.Md, A.Nm)
End Function

Function Mth_LnoAy(A As Mth) As Integer()
Mth_LnoAy = Ay_XAdd_1(SrcMthNm_MthIx(Md_Src(A.Md), A.Nm))
End Function

Function MdMthNm_FmCntAy(A As CodeModule, MthNm$) As FmCnt()
MdMthNm_FmCntAy = SrcMthNm_FmCntAy(Md_Src(A), MthNm)
End Function

Private Sub Z_MdMthNm_FmCntAy()
Dim A() As FmCnt: A = MdMthNm_FmCntAy(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FmCntDmp A(J)
Next
End Sub

Function Mth_MdDNm$(A As Mth)
Mth_MdDNm = Md_DNm(A.Md)
End Function

Function Mth_MdNm$(A As Mth)
Mth_MdNm = Md_Nm(A.Md)
End Function

Function Mth_PjNm$(A As Mth)
Mth_PjNm = Mth_PjNm(A.Md)
End Function

Function Mth_IsFun(A As Mth) As Boolean
Mth_IsFun = Md_IsStd(A.Md)
End Function

Function Mth_Exist(A As Mth) As Boolean
Mth_Exist = Md_XHas_MthNm(A.Md, A.Nm)
End Function


Sub Mth_XRpl(A As Mth, By$)
Mth_XRmv A
Md_XApp_Lines A.Md, By
End Sub

Function Mth_IsPub(A As Mth) As Boolean
Const CSub$ = CMod & "Mth_IsPub"
Dim L$: L = Mth_MthLin(A): If L = "" Then XThw CSub, "Mth does not have MthLin", "Mth", Mth_MthDNm(A)
Mth_IsPub = Lin_IsPubMth(L)
End Function


Private Sub Z()
Z_MdMthNm_FmCntAy
MIde_X_Mth:
End Sub
