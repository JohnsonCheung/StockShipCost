Attribute VB_Name = "MIde__Dft"
Option Compare Binary
Option Explicit
Function MdNm_DftMd(MdNm$) As CodeModule
If MdNm = "" Then
    Set MdNm_DftMd = CurMd
Else
    Set MdNm_DftMd = Md(MdNm)
End If
End Function
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CurMd
Else
   Set DftMd = A
End If
End Function

Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
   Set DftPj = CurPj
Else
   Set DftPj = A
End If
End Function


Function DftMdyAy(A$) As String()
DftMdyAy = CvNy(A)
End Function

Function DftMth(Mth_MthDNm0$) As Mth
If Mth_MthDNm0 = "" Then
    Set DftMth = CurMth
    Exit Function
End If
Set DftMth = DDNmMth(Mth_MthDNm0)
End Function
Function Pj_Sz&(A As VBProject)
Dim O&, I
For Each I In AyNz(Pj_MdAy(A))
    O = O + Md_Sz(CvMd(I))
Next
Pj_Sz = O
End Function
Function CurPj_Sz&()
CurPj_Sz = Pj_Sz(CurPj)
End Function
Function PjNm_DftPj(Pj_Nm$) As VBProject
If Pj_Nm = "" Then
    Set PjNm_DftPj = CurPj
Else
    Set PjNm_DftPj = Pj(Pj_Nm)
End If
End Function

Function DftFunByDDNm(MthDDNm0$) As Mth
If MthDDNm0 = "" Then
    Dim M As Mth
    Set M = CurMth
    If Mth_IsFun(M) Then
        Set DftFunByDDNm = M
    End If
Else
End If
Stop '
End Function
Function DftCmpTyAy(A) As vbext_ComponentType()
If IsLngAy(A) Then DftCmpTyAy = A
End Function

Private Sub Z_DftCmpTyAy()
Dim X() As vbext_ComponentType
DftCmpTyAy (X)
Stop
End Sub


Function DftMthNm$(MthNm0$)
If MthNm0 = "" Then
    DftMthNm = CurMthNm
    Exit Function
End If
DftMthNm = MthNm0
End Function




Private Sub Z()
Z_DftCmpTyAy
MIde__Dft:
End Sub
