Attribute VB_Name = "MIde_Loc_Go"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Loc_Go."
Sub Md_XGoLno(A As CodeModule, Lno&)
'Md_XGoRRCC A, RRCC(Lno, Lno, 1, 1)
End Sub

Sub MdLCC_XGo(A As CodeModule, B As LCC)
Const CSub$ = CMod & "MdLCC_XGo"
Md_XGo A
'If IsNothing(LCC) Then Msg CSub, "Given LCC is nothing": Exit Sub
With B
    A.CodePane.TopLine = .Lno
    A.CodePane.SetSelection .Lno, .C1, .Lno, .C2
End With
SendKeys "^{F4}"
End Sub
Sub MdMth_XGo(A As CodeModule, MthNm$)
'Md_XGoRRCC A, MdMthRRCC(A, MthNm)
End Sub

Function Mth_LCC(A As Mth) As LCC
Dim L%, C As LCC
Dim M As CodeModule
Set M = A.Md
For L = M.CountOfDeclarationLines + 1 To M.CountOfLines
    C = LinMthNm_LCC(M.Lines(L, 1), A.Nm, L)
    'If Not IsNothing(C) Then
    '    Set Mth_LCC = C
        Exit Function
    'End If
Next
End Function

Sub Mth_XGo(A As Mth)
'Md_XGoLCC A.Md, Mth_LCC(A)
End Sub

Sub Md_XGo1(A As CodeModule, B As VbeLoc)
'If IsEmpRRCC(RRCC) Then Debug.Print QQ_Fmt("Given RRCC_ is empty"): Exit Sub
Md_XShw A
'If IsRRCC_OutSidMd(RRCC, A) Then
'    With RRCC
'        Debug.Print QQ_Fmt("Md_XGoRg: Given ? is outside given Md[?]-(MaxR ?)(MaxR1C ?)(MaxR2C ?)", RRCC_Str(RRCC), Md_Nm(A), MdNLin(A), Len(A.Lines(.R1, 1)), Len(A.Lines(.R2, 1)))
'    End With
    Exit Sub
'End If
'With RRCC
'    A.CodePane.SetSelection .R1, .C1, .R2, .C2
'End With
End Sub

Sub Md_XGoTy(A As CodeModule, TyNm$)
'Md_XGoRRCC A, Md_TyRRCC(A, TyNm)
End Sub

Sub Lno_XGo(Lno&)
CurMdLno_XGo Lno
End Sub
Sub CurMdLno_XGo(Lno&)
MdLno_XGo CurMd, Lno
End Sub

Sub MdLno_XGo(A As CodeModule, Lno&)
Md_XGo A
Dim EndCol&
EndCol = Len(A.Lines(Lno, 1)) + 1
CurCdPne.SetSelection Lno, 1, Lno, EndCol
End Sub
Sub MdNmLno_XGo(MdNm, Lno&)
MdLno_XGo Md(MdNm), Lno
End Sub

Sub Md_XGo(A As CodeModule)
Md_XShw A
BrwObjWin.Visible = True
'WinApKeep Md_Win(A), BrwObjWin
XClr_ImmWin
XTile_V
End Sub
Sub Pj_XGo(A As VBProject)
ClsAllWin
Dim Md As CodeModule
Set Md = PjFstMbr(A)
If IsNothing(Md) Then
    Exit Sub
End If
Md.CodePane.Show
TileVBtn.Execute
DoEvents
End Sub

Sub PjMdNm_XGo(A As VBProject, MdNm$)
XCls_AllWin_Exl Ap_WinAy(Md_Win(Md(MdNm)))
End Sub
