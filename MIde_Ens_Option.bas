Attribute VB_Name = "MIde_Ens_Option"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Ens_Option."
Const OptExp$ = "Option Explicit"
Const OptCmpBin$ = "Option Compare Binary"
Const OptCmpDb$ = "Option Compare Database"
Const OptCmpTxt$ = "Option Compare Text"

Private Function Has_OptLin(A As CodeModule, XXX$) As Boolean
Dim I
For Each I In AyNz(Md_DclLy(A))
   If XHas_Pfx(I, XXX) Then Has_OptLin = True: Exit Function
Next
End Function

Private Function Md_OptCmpDbLno%(A As CodeModule)
Dim Ay$(): Ay = Md_DclLy(A)
Dim J%
For J = 0 To UB(Ay)
    If XHas_Pfx(Ay(J), OptCmpDb) Then Md_OptCmpDbLno = J + 1: Exit Function
Next
End Function

Private Function Md_OptCmpTxtLno%(A As CodeModule)
Dim Ay$(): Ay = Md_DclLy(A)
Dim J%
For J = 0 To UB(Ay)
    If XHas_Pfx(Ay(J), OptCmpTxt) Then Md_OptCmpTxtLno = J + 1: Exit Function
Next
End Function

Private Sub Md_XEns_OptCmpBinLin(A As CodeModule)
Md_XRmv_EmpLinBetTwoOpt A
If Has_OptLin(A, OptCmpBin) Then Exit Sub
Md_XRmv_OptCmpDbLin A
A.InsertLines 1, OptCmpBin
Debug.Print Md_Nm(A), , OptCmpBin; " is XEnsured ..."
End Sub

Private Sub Md_XEns_OptExpLin(A As CodeModule)
Const CSub$ = CMod & "Md_XEns_OptExpLin"
Md_XRmv_EmpLinBetTwoOpt A
If Has_OptLin(A, OptExp) Then Exit Sub
A.InsertLines 1, OptExp
FunMsgNyAp_XDmp CSub, "[" & OptExp & "] is XEnsusred", "Md", Md_Nm(A)
End Sub

Private Sub Md_XXRmv_OptCmpTxtLin(A As CodeModule)
Const CSub$ = CMod & "Md_XXRmv_OptCmpTxtLin"
Dim I%: I = Md_OptCmpTxtLno(A)
If I = 0 Then Exit Sub
A.DeleteLines I
FunMsgNyAp_XDmp CSub, "[Option Compare Text] line is deleted", "Md Lno", Md_Nm(A), I
End Sub

Private Sub Md_XRmv_EmpLinBetTwoOpt(A As CodeModule)
Const CSub$ = CMod & "Md_XRmv_EmpLinBetTwoOpt"
Const C = "Option"
If Not XHas_Pfx(A.Lines(1, 1), C) Then Exit Sub
If Trim(A.Lines(2, 1)) <> "" Then Exit Sub
If Not XHas_Pfx(A.Lines(3, 1), C) Then Exit Sub
A.DeleteLines 2, 1
Msg CSub, "Empty line between 2 option lines is removed (Md=" & Md_Nm(A) & ")"
End Sub

Private Sub Md_XRmv_OptCmpDbLin(A As CodeModule)
Const CSub$ = CMod & "Md_XRmv_OptCmpDbLin"
Dim I%: I = Md_OptCmpDbLno(A)
If I = 0 Then Exit Sub
A.DeleteLines I
FunMsgNyAp_XDmp CSub, "[Option Compare Database] line is deleted", "Md Lno", Md_Nm(A), I
End Sub

Sub Pj_XEns_OptCmpBinLin(A As VBProject)
Dim M
For Each M In Pj_MdAy(A)
    Md_XEns_OptCmpBinLin CvMd(M)
Next
End Sub

Private Sub Pj_XEns_OptExp(A As VBProject)
Dim M
For Each M In Pj_MdAy(A)
    Md_XEns_OptExpLin CvMd(M)
Next
End Sub

Sub Pj_XRmv_OptCmpDbLin(A As VBProject)
Dim I
For Each I In Pj_MdAy(A)
   Md_XRmv_OptCmpDbLin CvMd(I)
Next
End Sub

Sub XEns_OptCmpBin_MD(Optional MdNm$)
Md_XEns_OptCmpBinLin MdNm_DftMd(MdNm)
End Sub

Sub XEns_OptCmpBin_PJ(Optional PjNm$)
Pj_XEns_OptCmpBinLin PjNm_DftPj(PjNm)
End Sub

Sub XEns_OptExp_MD(Optional MdNm$)
Md_XEns_OptExpLin MdNm_DftMd(MdNm)
End Sub

Sub XEns_OptExp_PJ(Optional PjNm$)
Pj_XEns_OptExp PjNm_DftPj(PjNm)
End Sub

Sub XEns_OptExp_VBE()
Dim P As VBProject
For Each P In CurVbe.VBProjects
    Pj_XEns_OptExp P
Next
End Sub

Private Sub Z_XEns_OptExp_MD()
XEns_OptExp_MD
End Sub

Private Sub Z_XEns_OptExp_PJ()
XEns_OptExp_PJ
End Sub

Private Sub ZZ()
Dim A$
Dim B As VBProject
Dim XX
Pj_XEns_OptCmpBinLin B
Pj_XRmv_OptCmpDbLin B
XEns_OptCmpBin_MD A
XEns_OptCmpBin_PJ A
XEns_OptExp_MD A
XEns_OptExp_PJ A
XEns_OptExp_VBE
End Sub

Private Sub Z()
Z_XEns_OptExp_MD
Z_XEns_OptExp_PJ
End Sub
