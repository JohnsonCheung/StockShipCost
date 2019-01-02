Attribute VB_Name = "MIde_Mth_Nm_Dic"
Option Compare Binary
Option Explicit

Function Md_MthNmDic(A As CodeModule) As Dictionary
Set Md_MthNmDic = Src_MthNmDic(Md_Src(A))
End Function

Function Src_MthNmDic(A$()) As Dictionary
'Return a dic with Key=MthNm and Val=MthLines with merging of Prp-Get-Let-Set of one-Lines
Dim Ix, MthNm$, Lines$
Dim O As New Dictionary
O.Add "*Dcl", Src_DclLines(A)
For Each Ix In AyNz(Src_MthIxAy(A))
    Lines = SrcMthIx_MthLines_WithTopRmk(A, CLng(Ix))
    MthNm = Lin_MthNm(A(Ix)): If MthNm = "" Then Stop
    Dic_XIup O, MthNm, Lines, vbCrLf & vbCrLf
Next
Set Src_MthNmDic = O
End Function

Private Sub Z_SrcMthNmDic()
Dic_XBrw Src_MthNmDic(CurSrc)
End Sub
