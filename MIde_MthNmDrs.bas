Attribute VB_Name = "MIde_MthNmDrs"
Option Explicit
Option Compare Database

Function Src_MthNmDrs(A$()) As Drs
Dim Dry(), L
For Each L In AyNz(A)
    PushI_SomSz Dry, Lin_MthNmBrk(L)
Next
Set Src_MthNmDrs = New_Drs(MthNmFny, Dry)
End Function

Function MthNmFny() As String()
MthNmFny = Ssl_Sy("MthNm Kd Mdy")
End Function

Function CurSrc_MthNmDrs() As Drs
Set CurSrc_MthNmDrs = Src_MthNmDrs(CurSrc)
End Function

Sub CurSrc_MthNmDrs_XBrw()
XBrw CurSrc_MthNmDrs
End Sub

Sub CurMd_MthNmDrs_XBrw()
CurSrc_MthNmDrs_XBrw
End Sub

Function CurMd_MthNmDrs() As Drs
Set CurMd_MthNmDrs = CurSrc_MthNmDrs
End Function
Function Md_MthNmDry(A As CodeModule, Optional AddMdNm As Boolean) As Variant()
Dim O()
O = Src_MthNmDry(Md_Src(A))
If AddMdNm Then
    O = Dry_XAdd_Col(O, Md_Nm(A))
End If
Md_MthNmDry = O
End Function

Function Md_MthNmDrs(A As CodeModule, Optional AddMdNm As Boolean) As Drs
Dim O As Drs
Set O = Src_MthNmDrs(Md_Src(A))
If AddMdNm Then
    Set O = Drs_XAdd_ConstCol(O, "MdNm", Md_Nm(A))
End If
Set Md_MthNmDrs = O
End Function

Function Pj_MthNmDrs(A As VBProject, Optional AddPjNm As Boolean) As Drs
Dim O(), I
For Each I In AyNz(Pj_MdAy(A))
    PushIAy O, Md_MthNmDry(CvMd(I), AddMdNm:=True)
Next
Dim Fny$()
    Fny = MthNmFny
    PushI Fny, "MdNm"
If AddPjNm Then
    O = Dry_XAdd_ConstCol(O, Pj_Nm(A))
    Push Fny, "PjNm"
End If
Set Pj_MthNmDrs = New_Drs(Fny, O)
End Function

Function CurPj_MthNmDt() As Dt
Set CurPj_MthNmDt = Pj_MthNmDt(CurPj)
End Function
Function CurPj_MthNmWs_XBrw() As Worksheet
Set CurPj_MthNmWs_XBrw = Ws_XVis(MthNmWs_XFmt(Dt_Ws(Pj_MthNmDt(CurPj))))
End Function

Function Pj_MthNmDt(A As VBProject) As Dt
Dim DtNm$
    DtNm = "Pj-MthNm-" & Pj_Nm(A)
Set Pj_MthNmDt = Drs_Dt(Pj_MthNmDrs(A), DtNm)
End Function
Function CurPj_MthNmDrs(Optional AddPjNm As Boolean) As Drs
Set CurPj_MthNmDrs = Pj_MthNmDrs(CurPj, AddPjNm)
End Function

Sub CurPj_MthNmDrs_XBrw(Optional AddPjNm)
XBrw CurPj_MthNmDrs
End Sub

