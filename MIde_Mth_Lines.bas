Attribute VB_Name = "MIde_Mth_Lines"
Option Compare Binary
Option Explicit

Property Get CurMth_MthLines$()
CurMth_MthLines = MdMthNm_MthLines(CurMd, CurMthNm$)
End Property

Function MdMthNm_MthLines$(A As CodeModule, MthNm)
MdMthNm_MthLines = SrcMthNm_MthLines(Md_BdyLy(A), MthNm)
End Function

Function MdMthBdyLy(A As CodeModule, MthNm) As String()
MdMthBdyLy = SrcMthNm_MthLines(Md_Src(A), MthNm)
End Function

Function MdMthNm_Lines$(A As CodeModule, MthNm, Optional WithTopRmk As Boolean)
MdMthNm_Lines = SrcMthNm_MthLines(Md_Src(A), MthNm, WithTopRmk)
End Function

Function MdMthNm_LnoLines$(A As CodeModule, Mth_Lno&)
MdMthNm_LnoLines = A.Lines(Mth_Lno, MdMthLno_MthNLin(A, Mth_Lno))
End Function

Function MdMthNm_MthLy(A As CodeModule, MthNm) As String()
MdMthNm_MthLy = SrcMthNm_MthLines(Md_BdyLy(A), MthNm)
End Function

Function MthDDNmLines$(MthDDNm$)
MthDDNmLines = MthLines(DDNmMth(MthDDNm))
End Function

Function MthLin_MthEndLin$(A$)
Dim B$
B = Lin_MthKd(A): If B = "" Then XThw CSub, "Given Line does not have MthKd", "Lin", A
MthLin_MthEndLin = "End " & B
End Function

Function MthLinCnt%(A As Mth)
MthLinCnt = FmCntAyLinCnt(Mth_FmCnt(A))
End Function

Function MthLines$(A As Mth, Optional WithTopRmk As Boolean)
MthLines = SrcMthNm_MthLines(Md_BdyLy(A.Md), A.Nm, WithTopRmk)
End Function

Function MthLinesWithTopRmk$(A As Mth)
MthLinesWithTopRmk = MthLines(A, WithTopRmk:=True)
End Function

'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property


Private Sub Z_MthDDNmLines()
GoTo ZZ
ZZ:
Debug.Print MthDDNmLines("QIde.MIde_Mth_Lines.ZZ_MthDDNmLines")
End Sub

'aa
Private Sub Z_MthLines()
Debug.Print MthLines(New_Mth(CurMd, "XX1"), WithTopRmk:=True)
End Sub

Private Sub Z()
Z_MthDDNmLines
Z_MthLines
MIde_Mth_Lines:
End Sub
