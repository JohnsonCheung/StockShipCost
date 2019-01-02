Attribute VB_Name = "MIde_Mth_Lin_TyChr"
Option Compare Binary
Option Explicit
Const MthTyChrLis$ = "!@#$%^&"

Function IsMthTyChr(A$) As Boolean
If Len(A) <> 1 Then Exit Function
IsMthTyChr = HasSubStr(MthTyChrLis, A)
End Function


Function MthTyChrAsTyStr$(MthTyChr$)
Dim O$
Select Case MthTyChr
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Stop
End Select
MthTyChrAsTyStr = O
End Function

Function RmvMthTyChr$(A)
If IsMthTyChr(XTak_FstChr(A)) Then RmvMthTyChr = XRmv_FstChr(A) Else RmvMthTyChr = A
End Function

Function XShf_MthShtTyChr$(OLin)
XShf_MthShtTyChr = XShf_Chr(OLin, MthTyChrLis)
End Function
