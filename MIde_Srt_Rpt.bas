Attribute VB_Name = "MIde_Srt_Rpt"
Option Compare Binary
Option Explicit
Private Function Src_SrtRptLy(A$()) As String()
Dim A1 As Dictionary
Dim B1 As Dictionary
Set A1 = Src_MthNmDic(A)
Set B1 = Src_MthNmDic(Src_SrtLy(A))
Src_SrtRptLy = Dic_CmpFmt(A1, B1, "BefSrt", "AftSrt")
End Function

Private Sub Z_Src_SrtRptLy()
Brw Src_SrtRptLy(CurSrc)
End Sub

Property Get CurMd_SrtRptLy() As String()
CurMd_SrtRptLy = Md_SrtRptLy(CurMd)
End Property

Function Pj_SrtRptLy(A As VBProject) As String()
Dim O$(), I
For Each I In AyNz(Pj_MdAy(A))
    PushAy O, Md_SrtRptLy(CvMd(I))
Next
Pj_SrtRptLy = O
End Function

Function Pj_SrtRptDic(A As VBProject) As Dictionary 'Return a dic of [MdNm,SrtCmpFmt]
'SrtCmpDic is a New_Dic_LY with Key as MdNm and value is SrtCmpLy
Dim I, O As New Dictionary, Md As CodeModule
    For Each I In AyNz(Pj_MdAy(A))
        Set Md = I
        O.Add Md_Nm(Md), Md_SrtRptLy(Md)
    Next
Set Pj_SrtRptDic = O
End Function

Function Md_SrtRptLy(A As CodeModule) As String()
Md_SrtRptLy = Src_SrtRptLy(Md_Src(A))
End Function

Function Md_SrtDic(A As CodeModule) As Dictionary
Set Md_SrtDic = DicAddKeyPfx(Src_SrtDic(Md_Src(A)), Md_Nm(A) & ".")
End Function

Sub CurMd_SrtRpt_XBrw()
Brw Md_SrtRptLy(CurMd)
End Sub

