Attribute VB_Name = "MIde_Cmd_Qry"
Option Compare Binary
Option Explicit
Sub LisMth(Pfx$, Optional InclPrv As Boolean)
LisMth_PFX Pfx, InclPrv
End Sub

Sub LisMth_PFX(Pfx$, Optional InclPrv As Boolean)
Dim A$
    A = "Pub" & IIf(InclPrv, " Prv", "")
D Ay_XAlign_AtDot(AyQSrt(Pj_MthNy(CurPj, New_WhMdMth(Mth:=WhMth(A, Nm:=WhNm_PFX(Pfx))))))
End Sub

Sub LisMth_SFX(Sfx$, Optional InclPrv As Boolean)
Dim A$
    A = "Pub" & IIf(InclPrv, " Prv", "")
D Ay_XAlign_AtDot(AyQSrt(Pj_MthNy(CurPj, New_WhMdMth(Mth:=WhMth(A, Nm:=WhNm_SFX(Sfx))))))
End Sub

Sub LisMth_PATN(Patn$, Optional InclPrv As Boolean)
Dim A$
    A = "Pub" & IIf(InclPrv, " Prv", "")
D Ay_XAlign_AtDot(AyQSrt(Pj_MthNy(CurPj, New_WhMdMth(Mth:=WhMth(A, Nm:=WhNm(Patn))))))
End Sub

