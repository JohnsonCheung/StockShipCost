Attribute VB_Name = "MIde_Mth_Key"
Option Compare Binary
Option Explicit

Function Pj_MthKy(A As VBProject, Optional IsWrap As Boolean) As String()
Pj_MthKy = AyMapPX_Sy(Pj_MdAy(A), "MdMthKy", IsWrap)
End Function

Function PjMthKySq(A As VBProject) As Variant()
PjMthKySq = MthKy_Sq(Pj_MthKy(A, True))
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = Ws_XVis(Sq_Ws(PjMthKySq(A)))
End Function


