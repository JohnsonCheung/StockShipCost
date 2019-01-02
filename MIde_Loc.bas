Attribute VB_Name = "MIde_Loc"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Loc."
Function SrcMthNm_LCC(A$(), MthNm) As LCC
Dim R&, C&, Ix&
Ix = SrcMthNm_MthIx(A, MthNm)
R = Ix + 1
C = InStr(A(Ix), MthNm)
SrcMthNm_LCC = LCC(R, C + 1, C + Len(MthNm))
End Function


Function IsRRCC_OutSidMd(A As RRCC, Md As CodeModule) As Boolean
IsRRCC_OutSidMd = True
Dim R%
R = Md.CountOfLines
'If RRCCIsEmp(A) Then Exit Function
With A
   If .R1 > R Then Exit Function
   If .R2 > R Then Exit Function
   If .C1 > Len(Md.Lines(.R1, 1)) + 1 Then Exit Function
   If .C2 > Len(Md.Lines(.R2, 1)) + 1 Then Exit Function
End With
IsRRCC_OutSidMd = False
End Function

Sub LocStr_XGo(A)
Loc_XGo LocStr_Loc(A)
End Sub

Sub CdPneLCC_XGo(A As CodePane, B As LCC)
Dim C As LCC
    C = LCC_NormLCC(B)
With C
    A.SetSelection .Lno, .C1, .Lno, .C2
End With
End Sub

Function LCC_NormLCC(A As LCC) As LCC
With LCC_NormLCC
    If A.Lno <= 0 Then .Lno = 1 Else .Lno = A.Lno
    If A.C1 <= 0 Then .C1 = 1 Else .C1 = A.C1
    If A.C2 < .C1 Then .C2 = .C1 Else .C2 = A.C2
End With
End Function

Function LocStr_Loc(A) As VbeLoc
Dim Pj$, Md$, Lno&, C1%, C2%
With Brk(A, ":")
    With Brk2(.S1, ".")
        Pj = .S1
        Md = .S2
    End With
    Dim Ay$()
    Ay = SplitDot(.S2)
    Select Case Sz(Ay)
    Case 1: Lno = Ay(0)
    Case 2: Lno = Ay(0): C1 = Ay(1)
    Case 3: Lno = Ay(0): C1 = Ay(1): C2 = Ay(2)
    End Select
End With
Set LocStr_Loc = New_Loc(Pj, Md, Lno, C1, C2)
End Function

Function New_Loc(Pj$, Md$, Lno&, C1%, C2%) As VbeLoc
Set New_Loc = New VbeLoc
With New_Loc
    .Pj = Pj
    .Md = Md
    .Lno = Lno
    .C1 = C1
    .C2 = C2
End With
End Function

Function Loc_CdPne(A As VbeLoc) As CodePane
Set Loc_CdPne = Loc_Md(A).CodePane
End Function

Function Loc_Md(A As VbeLoc) As CodeModule
Set Loc_Md = Pj_Md(Pj(A.Pj), A.Md)
End Function

Sub Loc_XGo(A As VbeLoc)
CdPneLCC_XGo Loc_CdPne(A), Loc_LCC(A)
End Sub
Function Loc_LCC(A As VbeLoc) As LCC
With Loc_LCC
    .Lno = A.Lno
    .C1 = A.C1
    .C2 = A.C2
End With
End Function
Private Sub Z_Mth_FmCnt()
Dim M As Mth: Set M = New_Mth(Md("ZZModule"), "YYA")
Dim Act() As FmCnt: Act = Mth_FmCnt(M)
Ass Sz(Act) = 2
Ass Act(0).FmLno = 5
Ass Act(0).Cnt = 7
Ass Act(1).FmLno = 13
Ass Act(1).Cnt = 15
End Sub

Function LinMthNm_LCC(A$, MthNm$, Lno%) As LCC
Const CSub$ = CMod & "LinMthNm_LCC"
Dim M$: M = Lin_MthNm(A):
If M = "" Then
    XThw CSub, "Given Lin is not a MthLin", _
    "Given-Lin MthNm Lno", A, MthNm, Lno
End If
If M <> MthNm Then
    XThw CSub, "Given Lin does not have MthNm", "Lin MthNm Lno", A, MthNm, Lno
End If
Dim C1%, C2%
C1 = InStr(A, MthNm)
C2 = C1 + Len(MthNm)
LinMthNm_LCC = LCC(Lno, C1, C2)
End Function

Function MdMthNm_Loc(A As CodeModule, MthNm$) As VbeLoc
'MdMthLoc = SrcMthRRCC(Md_Src(A), MthNm)
End Function

Function RRCC_Str$(A As RRCC)
With A
'   RRCC_Str = QQ_Fmt("(RRCC : ? ? ? ??)", .R1, .R2, .C1, .C2, IIf(IsEmpRRCC(A), " *Empty", ""))
End With
End Function


Private Sub Z()
Z_Mth_FmCnt
MIde_Loc:
End Sub
