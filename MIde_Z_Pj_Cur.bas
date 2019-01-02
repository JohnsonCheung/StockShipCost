Attribute VB_Name = "MIde_Z_Pj_Cur"
Option Compare Binary
Option Explicit

Property Get CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Property

Sub CurPj_XAdd_Md(Nm$)
Pj_XAdd_Md CurPj, Nm
End Sub

Sub CurPj_XDlt_Md(MdNm$)
Pj_XDlt_Md CurPj, MdNm
End Sub

Function CurPj_XEns_Md(MdNm$) As CodeModule
Set CurPj_XEns_Md = Pj_XEns_Md(CurPj, MdNm)
End Function
Property Get PjFfnAy() As String()
PjFfnAy = CurVbe_PjFfnAy
End Property
Property Get CurVbe_PjFfnAy() As String()
PushIAy CurVbe_PjFfnAy, Vbe_PjFfnAy(CurVbe)
End Property

Property Get CurPj_FunPfxAy() As String()
CurPj_FunPfxAy = Pj_FunPfxAy(CurPj)
End Property

Function CurPj_MdAy(Optional A As WhMd) As CodeModule()
CurPj_MdAy = Pj_MdAy(CurPj, A)
End Function

Property Get CurPj_Nm$()
CurPj_Nm = CurPj.Name
End Property

Property Get CurPj_Pth$()
CurPj_Pth = Pj_Pth(CurPj)
End Property
