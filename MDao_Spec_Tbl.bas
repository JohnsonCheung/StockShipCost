Attribute VB_Name = "MDao_Spec_Tbl"
Option Compare Binary
Option Explicit
Property Get Spnm_Fn$(A)
Spnm_Fn = A & ".txt"
End Property

Property Get Spnm_Fny() As String()
Spnm_Fny = Dbt_Fny(CurDb, "Spec")
End Property

Function Spnm_Ft$(A)
Spnm_Ft = SpecPth & Spnm_Fn(A)
End Function

Sub Spnm_XImp(A)
Db_XImp_Spec CurDb, A
End Sub

Sub Spnm_XKill(A)
Kill Spnm_Ft(A)
End Sub

Property Get Spnm_Lines$(A)
Spnm_Lines = Spnm_V(A, "Lines")
End Property

Property Let Spnm_Lines(A, Lines$)
TfkV("Spec", A, "Lines") = Lines
End Property

Function Spnm_Ly(A) As String()
Spnm_Ly = SplitCrLf(Spnm_Lines(A))
End Function

Function Spnm_Tim(A) As Date
Spnm_Tim = Nz(TfkV("Spec", "Tim", A), 0)
End Function

Property Get Spnm_V(A, ValNm$)
Spnm_V = TfkV("Spec", ValNm, A)
End Property

Sub Spnm_XBrw(A)
Ft_XBrw Spnm_Ft(A)
End Sub

Sub Spnm_XEdt(A)
Spnm_XEnsRec A
Spnm_XExp_IfNotExist A
Spnm_XBrw A
End Sub

Sub Spnm_XEnsRec(A)
If Tbl_XHas_SkValAp("Spec", A) Then Exit Sub
TblValAp_XIns "Spec", A
End Sub

Sub Spnm_XExp(A)
Str_XWrt Spnm_Lines(A), Spnm_Ft(A), OvrWrt:=True
End Sub

Sub Spnm_XExp_IfNotExist(A)
If Ffn_Exist(A) Then Exit Sub
Str_XWrt Spnm_Lines(A), Spnm_Ft(A)
End Sub
Sub FmtSpec_XExp()
Spnm_XExp FmtSpecNm
End Sub

Sub FmtSpec_XImp()
Spnm_XImp FmtSpecNm
End Sub

Sub FmtSpec_Kill()
Spnm_XKill FmtSpecNm
End Sub

Property Get FmtSpec_Lines$()
FmtSpec_Lines = Spnm_Lines(FmtSpecNm)
End Property

Property Get FmtSpec_Ly() As String()
FmtSpec_Ly = SplitCrLf(FmtSpec_Lines)
End Property

Sub FmtSpec_XBrw()
Spnm_XBrw FmtSpecNm
End Sub

Property Get FmtSpecErLy() As String()
Dim A$(), B$(), C$(), D$()
'Dim Fml$(), Wdt$()
'A = LyChk(FmtSpec_Ly, VdtFmtSpecNmSsl)
'B = XFmlChk(Fml)
'C = XWdtChk(Wdt)
'D = FmtChk(Fmt)
'E = XAlignC(AlignC)
'F = XTSumChk(TSum)
'G = XTAvgChk(TAvg)
'H = XTCntChk(TCnt)
'I = XReSeqChk(ReSeq)
End Property
