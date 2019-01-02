Attribute VB_Name = "MAcs_Acs_Cpy"
Option Compare Binary
Option Explicit
Sub Acs_XCpy_Rf(A As Access.Application, Fb$)

End Sub
Sub Acs_XCpy_Frm(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeProject.AllForms
    A.DoCmd.CopyObject Fb, , acForm, I.Name
Next
End Sub

Sub Acs_XCpy_Md(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeProject.AllModules
    A.DoCmd.CopyObject Fb, , acModule, I.Name
Next
End Sub

Sub Acs_XCpy_Tbl(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeData.AllTables
    If Not Tbl_IsSys(I.Name) Then
        A.DoCmd.CopyObject Fb, , acTable, I.Name
    End If
Next
End Sub

Sub Acs_XCpy(A As Access.Application, Optional Fb0$)
Dim Fb$
If Fb0 = "" Then
    Fb = Ffn_NxtFfn(A.CurrentDb.Name)
Else
    Fb = Fb0
End If
Ass PthIsExist(Ffn_Pth(Fb))
Ass Not Ffn_Exist(Fb)
Fb_XCrt Fb
Acs_XCpy_Tbl A, Fb
Acs_XCpy_Frm A, Fb
Acs_XCpy_Md A, Fb
Acs_XCpy_Rf A, Fb
End Sub

Sub CurAcs_XCpy(Optional Fb0$)
Acs_XCpy CurAcs, Fb0
End Sub
