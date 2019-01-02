Attribute VB_Name = "MIde_Z_Vbe_Cur"
Option Compare Binary
Option Explicit

Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Sub CurVbe_XExp()
Vbe_XExp CurVbe
End Sub

Function CurVbe_XHas_Bar(Nm$) As Boolean
CurVbe_XHas_Bar = Vbe_XHas_Bar(CurVbe, Nm)
End Function

Function CurVbe_XHas_Pj_Ffn(Pj_Ffn) As Boolean
CurVbe_XHas_Pj_Ffn = Vbe_XHas_Pj_Ffn(CurVbe, Pj_Ffn)
End Function


Function CurVbePj(A$) As VBProject
Set CurVbePj = CurVbe.VBProjects(A)
End Function

Function CurVbePj_FfnPj(Pj_Ffn) As VBProject
Set CurVbePj_FfnPj = VbePj_FfnPj(CurVbe, Pj_Ffn)
End Function

Sub CurVbePj_MdFmtBrw()
Brw VbePj_MdFmt(CurVbe)
End Sub

Sub CurVbeSav()
VbeSav CurVbe
End Sub

Property Get CurVbe_Src() As String()
CurVbe_Src = Vbe_Src(CurVbe)
End Property
