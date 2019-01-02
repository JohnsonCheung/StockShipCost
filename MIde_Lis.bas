Attribute VB_Name = "MIde_Lis"
Option Compare Binary
Option Explicit
Function Pj_MdLisDt(A As VBProject, Optional B As WhMd) As Dt
Stop '
End Function

Sub Pj_MdLisDt_XBrw(A As VBProject, Optional B As WhMd)
Dt_XBrw Pj_MdLisDt(A, B)
End Sub

Sub Pj_MdLisDt_XDmp(A As VBProject, Optional B As WhMd)
Dt_XDmp Pj_MdLisDt(A, B)
End Sub

Sub XLis_Md(Optional Patn$, Optional ExlLikss$)
D Ay_XSrt(Ay_XWh_PatnExl(Pj_MdNy(CurPj), Patn, ExlLikss))
End Sub
Sub XLis_Pj()
Dim A$()
    A = Vbe_PjNy(CurVbe)
    D Ay_XAdd_Pfx(A, "ShwPj """)
D A
End Sub
Function New_WhMdMth_MTH_MD(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$) As WhMdMth

End Function

Sub XLis_Mth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$)
Dim Ny$(), M As WhMdMth
    Set M = New_WhMdMth_MTH_MD(MthPatn, MthExl, WhMdy, WhKd, MdPatn)
    Ny = Pj_MthDDNy(CurPj, M)
D Ay_XAdd_Pfx(Ny, CurPj_Nm & ".")
End Sub

Property Get MdLisFny() As String()
MdLisFny = SplitSpc("PJ Md-Pfx Md Ty Lines NMth NMth-Pub NMth-Prv NTy NTy-Pub NTy-Prv NEnm NEnm-Pub NEnm-Prv")
End Property

