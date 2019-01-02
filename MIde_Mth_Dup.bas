Attribute VB_Name = "MIde_Mth_Dup"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Mth_Dup."
Function DupMthFNy_GpAy(A$()) As Variant()
Dim O(), J%, M$()
Dim L$ ' LasMthNm
L = Brk(A(0), ":").S1
Push M, A(0)
Dim B As S1S2
For J = 1 To UB(A)
    Set B = Brk(A(J), ":")
    If L <> B.S1 Then
        Push O, M
        Erase M
        L = B.S1
    End If
    Push M, A(J)
Next
If Sz(M) > 0 Then
    Push O, M
End If
DupMthFNy_GpAy = O
End Function

Function DupMthFNy_SamMthBdyFunFNy(A$(), Vbe As Vbe) As String()
Dim Gp(): Gp = DupMthFNy_GpAy(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupMthFNyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
DupMthFNy_SamMthBdyFunFNy = O
End Function

Sub DupMthFNy_ShwNotDupMsg(A$(), MthNm)
Select Case Sz(A)
Case 0: Debug.Print QQ_Fmt("DupMthFNy_ShwNotDupMsg: no such Fun(?) in CurVbe", MthNm)
Case 1
    Dim B As S1S2: Set B = Brk(A(0), ":")
    If B.S1 <> MthNm Then Stop
    Debug.Print QQ_Fmt("DupMthFNy_ShwNotDupMsg: Fun(?) in Md(?) does not have dup-Fun", B.S1, B.S2)
End Select
End Sub



Function DupMthFNyGp_IsDup(Ny) As Boolean
DupMthFNyGp_IsDup = Ay_IsAllEleEq(AyMap(Ny, "FunFNm_MthLines"))
End Function

Function DupMthFNyGp_IsVdt(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupMthFNyGp_IsVdt = True
End Function

Function DupMthFNyGpAy_AllSameCnt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupMthFNyGp_IsDup(Gp) Then O = O + 1
Next
DupMthFNyGpAy_AllSameCnt = O
End Function
Sub AAA()
Z_Pj_DupMthNy_WithLinesId
End Sub
Private Sub Z_Pj_DupMthNy_WithLinesId()
D Pj_DupMthNy_WithLinesId(CurPj)
End Sub

Function Pj_DupMthNy_WithLinesId(A As VBProject) As String()
Dim Dic As New Dictionary, N
For Each N In AyNz(PjDupMthNy(A))
    PushI Pj_DupMthNy_WithLinesId, N & "." & X1(A, N, Dic)
Next
End Function

Private Function X1%(Pj As VBProject, MdDotMthNm, Dic As Dictionary)
Dim Lines$, MdNm$, M As Mth, MthNm
BrkAsg MdDotMthNm, ".", MdNm, MthNm
Set M = New_Mth(Pj_Md(Pj, MdNm), MthNm)
Lines = MthLines(M, WithTopRmk:=True)
If Dic.Exists(Lines) Then X1 = Dic(Lines): Exit Function
Dim Ix%: Ix = Dic.Count
Dic.Add Lines, Ix
X1 = Ix
End Function

Private Sub Z_PjDupMthNy()
D PjDupMthNy(CurPj)
End Sub

Function PjDupMthNy(A As VBProject) As String()
Dim Dry()
Dry = PjDupMth_Pj_Md_Mth_Dry(A) ' Pj_Nm MdNm MthNm
Dry = Dry_XSrt(Dry, 2)
Dry = DrySelIxAp(Dry, 1, 2) ' MthNm MdNm
PjDupMthNy = DryMAp_JnDot(Dry)
End Function

Function VbeDupMdNy(A As Vbe) As String()
VbeDupMdNy = Dry_Fmtss(Dry_XWh_Dup(VbePj_MdDry(A)))
End Function

Function MthNm_DupMthFNy(A) As String()
Stop '
'MthNm_DupMthFNy = VbeFunFNm(CurVbe, FunPatn:="^" & A & "$")
End Function


Private Sub Z()
Z_PjDupMthNy
MIde_Mth_Dup:
End Sub
