Attribute VB_Name = "MIde_Mth_Cnt"
Option Compare Binary
Option Explicit

Function Md_NMth%(A As CodeModule, Optional B As WhMth)
Md_NMth = Src_NMth(Md_Src(A), B)
End Function

Function MdMthLno_MthNLin%(A As CodeModule, MthLno&)
Dim Kd$, Lin$, EndLin$, J%
Lin = A.Lines(MthLno, 1)
Kd = Lin_MthKd(Lin)
If Kd = "" Then Stop
EndLin = "End " & Kd
If XHas_Sfx(Lin, EndLin) Then
    MdMthLno_MthNLin = 1
    Exit Function
End If
For J = MthLno + 1 To A.CountOfLines
    If XHas_Sfx(A.Lines(J, 1), EndLin) Then
        MdMthLno_MthNLin = J - MthLno + 1
        Exit Function
    End If
Next
Stop
End Function
Function Pj_NSrcLin&(A As VBProject)
Dim O&, C As VBComponent
For Each C In A.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
Pj_NSrcLin = O
End Function
Function Md_NMthPfx%(A As CodeModule)
Md_NMthPfx = Sz(Md_MthPfxAy(A))
End Function

Function Md_NMth_Pub%(A As CodeModule)
Md_NMth_Pub = Src_NMth(Md_Src(A), WhMth("Pub"))
End Function
Property Get CurPj_NMth_Pub%()
CurPj_NMth_Pub = Pj_NMth_Pub(CurPj)
End Property

Function Pj_NMth_Pub%(A As VBProject)
Dim O%, C As VBComponent
For Each C In A.VBComponents
    O = O + Md_NMth_Pub(C.CodeModule)
Next
Pj_NMth_Pub = O
End Function

Function Src_NMth%(A$(), Optional B As WhMth)
Src_NMth = Sz(Src_MthIxAy(A, B))
End Function
