Attribute VB_Name = "MIde_Z_Md_Emp"
Option Compare Binary
Option Explicit
Private Sub Z_Md_IsEmp()
Dim M As CodeModule
'GoSub Tst1
GoSub Tst2
Exit Sub
Tst2:
    Set M = Md("Module2")
    Ept = True
    GoSub Tst
    Return
Tst1:
    '-----
    Dim T$, P As VBProject
        Set P = CurPj
        T = TmpNm
    '---
    Set M = Pj_XAdd_Md(P, T)
    Ept = True
    GoSub Tst
    Pj_XDlt_Md P, T
    Return
Tst:
    Act = Md_IsEmp(M)
    C
    Return
End Sub

Function Src_IsEmp(A$()) As Boolean
Dim L
For Each L In AyNz(A)
    If Not Lin_IsEmpSrc(L) Then Exit Function
Next
Src_IsEmp = True
End Function

Function Md_IsEmp(A As CodeModule) As Boolean
Dim J%
For J = 1 To A.CountOfLines
    If Not Lin_IsEmpSrc(A.Lines(J, 1)) Then Exit Function
Next
Md_IsEmp = True
End Function

Function Lin_IsEmpSrc(A) As Boolean
Lin_IsEmpSrc = True
If XHas_Pfx(A, "Option ") Then Exit Function
Dim L$: L = Trim(A)
If L = "" Then Exit Function
If L = "'" Then Exit Function
Lin_IsEmpSrc = False
End Function

Private Sub Z_Pj_MdNy_EMP()
Brw Pj_MdNy_EMP(CurPj)
End Sub

Function Pj_MdNy_EMP(A As VBProject) As String()
Dim C As VBComponent, N$
N = A.Name & "."
For Each C In A.VBComponents
    Select Case C.Type
    Case vbext_ct_ClassModule, vbext_ct_StdModule
        If Md_IsEmp(C.CodeModule) Then PushI Pj_MdNy_EMP, N & C.Name & ":" & CmpTy_ShtNm(C.Type)
    End Select
Next
End Function

Function Vbe_MdNy_EMP(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy Vbe_MdNy_EMP, Pj_MdNy_EMP(P)
Next
End Function

Private Sub Z_Vbe_MdNy_EMP()
D Vbe_MdNy_EMP(CurVbe)
End Sub


Private Sub Z()
Z_Md_IsEmp
Z_Pj_MdNy_EMP
Z_Vbe_MdNy_EMP
MIde_Z_Md_Emp:
End Sub
