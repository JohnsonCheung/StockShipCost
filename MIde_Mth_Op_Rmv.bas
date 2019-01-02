Attribute VB_Name = "MIde_Mth_Op_Rmv"
Option Compare Binary
Option Explicit

Sub MdMthNm_XRmv(A As CodeModule, M)
Dim X() As FmCnt: X = MdMdMthNm_FmCntAyWithTopRmk(A, M)
XDmp_Ly CSub, "Remove method", "Md Mth FmCnt-WithTopRmk", Md_Nm(A), M, FmCntAyLy(X)
Md_XRmv_FmCntAy A, X
End Sub

Private Sub Z_Mth_XRmv()
Const N$ = "ZZModule"
Dim M As CodeModule
Dim M1 As Mth, M2 As Mth
GoSub Crt
Set M = Md(N)
Set M1 = New_Mth(M, "ZZRmv1")
Set M2 = New_Mth(M, "ZZRmv2")
Mth_XRmv M1
Mth_XRmv M2
Md_XRmv_EndBlankLin M
If M.CountOfLines <> 0 Then MsgBox M.CountOfLines
MdDlt M
Exit Sub
Crt:
    CurPj_XDlt_Md N
    Set M = CurPj_XEns_Md(N)
    Md_XApp_Lines M, XRpl_VBar("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Sub Mth_XRmv(A As Mth)
MdMthNm_XRmv A.Md, A.Nm
End Sub


Private Sub Z()
Z_Mth_XRmv
MIde_Mth_Op_Rmv:
End Sub
