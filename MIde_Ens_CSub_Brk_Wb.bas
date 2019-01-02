Attribute VB_Name = "MIde_Ens_CSub_Brk_Wb"
Option Compare Binary
Option Explicit

Function PjCSubBrkWb(A As VBProject) As Workbook
Set PjCSubBrkWb = Ds_Wb(ZDs(A))
End Function

Private Function ZAy1(A() As CSubBrk) As CSubBrkMd()
Dim I, M As CSubBrk
For Each I In AyNz(A)
    Set M = I
    PushObj ZAy1, M.MdBrk
Next
End Function

Private Function ZAy2(A() As CSubBrk) As CSubBrkMd()
Dim I, M As CSubBrk
For Each I In AyNz(A)
    Set M = I
    PushObjAy ZAy2, M.MthBrkAy
Next
End Function

Private Function ZDrs1(A() As CSubBrk) As Drs
Set ZDrs1 = Oy_Drs(ZAy1(A), ZFny1)
End Function

Private Function ZDrs2(A() As CSubBrk) As Drs
Set ZDrs2 = Oy_Drs(ZAy2(A), ZFny2)
End Function

Private Function ZDs(A As VBProject) As Ds
Dim DsNm$: DsNm = QQ_Fmt("CSubBrk:[Pj_Nm=?] [Pj_Ffn=?]", A.Name, Pj_Ffn(A))
Set ZDs = Ds(ZDtAy(A), DsNm)
End Function

Private Function ZDtAy(A As VBProject) As Dt()
Dim Ay() As CSubBrk
    Ay = PjCSubBrkAy(A)
PushObj ZDtAy, Drs_Dt(ZDrs1(Ay), "MdBrk")
PushObj ZDtAy, Drs_Dt(ZDrs2(Ay), "MthBrk")
End Function

Private Property Get ZFny1() As String()
Dim X As New CSubBrkMd
ZFny1 = Ssl_Sy(X.Fldss)
End Property

Private Property Get ZFny2() As String()
Dim X As New CSubBrkMth
ZFny2 = Ssl_Sy(X.Fldss)
End Property

Private Sub Z_PjCSubBrkWb()
Dim A As VBProject
GoTo ZZ
ZZ:
    Wb_XVis PjCSubBrkWb(CurPj)
End Sub


Private Sub Z()
Z_PjCSubBrkWb
MIde_XEns_CSub_Brk_Wb:
End Sub
