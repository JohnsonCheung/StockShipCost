Attribute VB_Name = "MIde_Mth_Dic"
Option Compare Binary
Option Explicit
Function Md_MthDic(A As CodeModule) As Dictionary
Set Md_MthDic = Md_MthDDNmDic(A)
End Function

Private Sub ZZ_Pj_MthDic()
Dic_XBrw Pj_MthDic(CurPj)
End Sub

Function Pj_MthDic(A As VBProject) As Dictionary
Dim I
Set Pj_MthDic = New Dictionary
For Each I In Pj_MdAy(A)
    PushDic Pj_MthDic, Md_MthDic(CvMd(I))
Next
End Function

Private Sub ZZ_MdMthDic()
Dic_XBrw Md_MthDic(CurMd)
End Sub


Private Sub Z_Md_MthDic()
Dic_XBrw Md_MthDic(CurMd)
End Sub
Private Sub Z_Pj_MthDic()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CurPj)
Ass IsSyDic(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub

Private Sub Z_Pj_MthDic1()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CurPj)
Ass IsSyDic(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub
Function PjFunBdyDic(A As VBProject) As Dictionary
Stop '
End Function

Private Sub Z()
End Sub
