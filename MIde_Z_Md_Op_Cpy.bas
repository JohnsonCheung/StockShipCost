Attribute VB_Name = "MIde_Z_Md_Op_Cpy"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Z_Md_Op_Cpy."

Private Sub Cmp_XCpy1(A As VBComponent, ToPj As VBProject)
Dim T$: T = TmpFt(Fnn:=A.Name)
A.Export T
ToPj.VBComponents.Import T
Kill T
End Sub

Sub CmpAy_XCpy(A() As VBComponent, ToPj As VBProject)
Dim I
For Each I In AyNz(A)
    Cmp_XCpy CvCmp(I), ToPj
Next
End Sub
Sub Cmp_XCpy(A As VBComponent, ToPj As VBProject)
Const CSub$ = CMod & "Cmp_XCpy"
Dim N$: N = A.Name
If Pj_XHas_CmpNm(ToPj, N) Then
    XThw CSub, "Cmp already exist", "CmpNm Cmp_PjNm TarPj", N, Cmp_PjNm(A), ToPj.Name
End If
If Cmp_IsCls(A) Then
    Cmp_XCpy1 A, ToPj 'If ClassModule need to export and import due to the Public/Private class property can only the set by Export/Import
Else
    Pj_XAdd_CmpLines ToPj, N, A.Type, Lines_XTrim_End(Md_Lines(A.CodeModule))
End If
Md_XCls_Win Pj_Md(ToPj, N)
If Trc Then Debug.Print QQ_Fmt("Cmp_XCpy: Cmp(?) is copied from SrcPj(?) to TarPj(?).", A.Name, Cmp_PjNm(A), ToPj.Name)
End Sub

Sub Md_XCpy(A As CodeModule, ToPj As VBProject)
Cmp_XCpy A.Parent, ToPj
End Sub

Private Sub ZZ()
Dim A As VBComponent
Dim B As VBProject
Dim D As CodeModule
Cmp_XCpy A, B
Md_XCpy D, B
End Sub

Private Sub Z()
End Sub
