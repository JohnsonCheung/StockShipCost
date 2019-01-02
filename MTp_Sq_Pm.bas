Attribute VB_Name = "MTp_Sq_Pm"
Option Compare Binary
Option Explicit
Private Function WLyRslt(A() As Lnx) As LyRslt
Dim O As LyRslt
Dim R1 As LnxAyRslt
Dim R2 As LnxAyRslt
    'ErIx1 = LnxAy_LnxAyRslt_DUP_KEY(A)
    'ErIx2 = LnxAy_LnxAyRslt_PERCENTAGE_PFX(A)

WLyRslt = O
End Function

Function PmBlkAy_PmRslt(A() As Lnx) As PmRslt
Dim O As PmRslt, LyR As LyRslt
LyR = WLyRslt(A)
    O.Er = LyR.Er
Set O.Pm = New_Dic_LINES(LyR.Ly)
PmBlkAy_PmRslt = O
End Function
