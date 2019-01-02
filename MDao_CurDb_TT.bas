Attribute VB_Name = "MDao_CurDb_TT"
Option Compare Binary
Option Explicit
Sub TT_XRen_ByXAdd_Pfx(TT$, Pfx$)
Dbtt_XRen_ByXAdd_Pfx CurDb, TT, Pfx
End Sub

Sub TT_XBrw(TT$)
Dbtt_Brw CurDb, TT
End Sub

Sub TT_XDrp(TT$)
Dbtt_XDrp CurDb, TT
End Sub

Sub TT_XOpn(TT)
Ay_XDo Ssl_Sy(TT), "TblOpn"
End Sub

Function TT_SclLy(TT) As String()
TT_SclLy = Ay_Sy(AyOfAy_Ay(AyMap(CvNy(TT), "Tbl_SclLy")))
End Function
