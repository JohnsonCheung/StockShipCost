Attribute VB_Name = "MIde_Z_Md_Op_Rmv_Lines"
Option Compare Binary
Option Explicit
Sub Md_XRmv_Bdy(A As CodeModule)
Md_XRmv_FmCnt A, MdBdyFmCnt(A)
End Sub

Sub Md_XRmv_Dcl(A As CodeModule)
If A.CountOfDeclarationLines = 0 Then Exit Sub
A.DeleteLines 1, A.CountOfDeclarationLines
End Sub


Sub Md_XClr(A As CodeModule, Optional IsSilent As Boolean)
With A
    If .CountOfLines = 0 Then Exit Sub
    If Not IsSilent Then Debug.Print QQ_Fmt("Md_XClr: Md(?) of lines(?) is cleared", Md_Nm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub Md_XClr_Bdy(A As CodeModule, Optional IsSilent As Boolean)
Stop
With A
    If .CountOfLines = 0 Then Exit Sub
    Dim N%, Lno%
        Lno = Md_BdyFmLno(A)
        N = .CountOfLines - Lno + 1
    If N > 0 Then
        If Not IsSilent Then Debug.Print QQ_Fmt("Md_XClr_Bdy: Md(?) of lines(?) from Lno(?) is cleared", Md_Nm(A), N, Lno)
        .DeleteLines Lno, N
    End If
End With
End Sub

Sub Md_XRmv_LNO(A As CodeModule, Lno)
If Lno = 0 Then Exit Sub
MsgAp_XDmp_Ly "Md_XRmv_LNO: [Md]-[Lno]-[Lin] is removed", Md_Nm(A), Lno, A.Lines(Lno, 1)
A.DeleteLines Lno, 1
End Sub

Sub Md_XRmv_EndBlankLin(A As CodeModule)
Dim J%
While A.CountOfLines > 1
    J = J + 1
    If J > 10000 Then Stop
    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
    A.DeleteLines A.CountOfLines, 1
Wend
End Sub

Sub Md_XDlt_FmCnt(A As CodeModule, B() As FmCnt)
If Not FmCntAy_IsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub
