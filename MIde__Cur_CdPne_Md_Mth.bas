Attribute VB_Name = "MIde__Cur_CdPne_Md_Mth"
Option Compare Binary
Option Explicit
Property Get CurMd_Nm$()
CurMd_Nm = CurCmp.Name
End Property

Function Md_CurMthNm$(A As CodeModule)
Dim R1&, R2&, C1&, C2&
A.CodePane.GetSelection R1, C1, R2, C2
Dim K As vbext_ProcKind
Md_CurMthNm = A.ProcOfLine(R1, K)
End Function

Property Get CurMd() As CodeModule
Set CurMd = CurCdPne.CodeModule
End Property

Private Sub Z_CurMd()
Ass CurMd.Parent.Name = "Cur_d"
End Sub

Property Get CurMd_Win() As VBIDE.Window
Dim A As CodePane
Set A = CurCdPne
If IsNothing(A) Then Exit Property
Set CurMd_Win = A.Window
End Property

Property Get CurMth() As Mth
Dim M As CodeModule
    Set M = CurMd
Set CurMth = New_Mth(M, Md_CurMthNm(M))
End Property

Property Get CurMthNm$()
CurMthNm = CurMth.Nm
End Property

Property Get CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Property

Private Sub Z()
Z_CurMd
MIde__Cur_CdPne_Md_Mth:
End Sub
