Attribute VB_Name = "MVb_X_WhNm"
Option Compare Binary
Option Explicit

Function WhNm(Optional Patn$, Optional Exl) As WhNm
Dim O As New WhNm
Set O.Re = New_Re(Patn)
O.ExlAy = CvSy(Exl)
Set WhNm = O
End Function

Function WhNm_PFX(Pfx$) As WhNm
Set WhNm_PFX = WhNm("^" & Pfx)
End Function

Function WhNm_SFX(Sfx$) As WhNm
Set WhNm_SFX = WhNm(Sfx & "$")
End Function

