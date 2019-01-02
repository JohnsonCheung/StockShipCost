Attribute VB_Name = "MIde_Mth_Op_Mov"
Option Compare Binary
Option Explicit
Sub CurMd_XMov_Mth(MthPatn$, ToMd As CodeModule)
MdMthNm_XMov CurMd, MthPatn, ToMd
End Sub

Sub XMov_Mth(MthPatn$, ToMdNm$)
CurMd_XMov_Mth MthPatn, Md(ToMdNm)
End Sub
Sub CurMth_XMov(ToMd$)
Mth_XMov CurMth, Md(ToMd)
End Sub

Sub MdMthNm_XMov(A As CodeModule, M$, ToMd As CodeModule)
If Md_XHas_MthNm(ToMd, M) Then XThw CSub, "Mth from Md already exist in ToMd", "MthNm FmMd ToMd", M, Md_Nm(ToMd), Md_Nm(A)
Md_XApp_Lines ToMd, MdMthNm_Lines(A, M)
MdMthNm_XRmv A, M
End Sub

Private Sub Z_MdMthNm_XMov()
'MdMthNm_XMov Md("Mth_"), "XX", Md("A_")
End Sub

Sub Mth_XMov(A As Mth, ToMd As CodeModule)
Const Trc As Boolean = False
If Trc Then
Debug.Print "Mth_XMov: Start........................\\"
Debug.Print "Mth_XMov: From :" & Mth_MthDNm(A)
Debug.Print "Mth_XMov: To   :" & Md_DNm(ToMd) & "." & A.Nm
End If
If Mth_XCpy(A, ToMd) Then
    Debug.Print "Mth_XMov: Fail to copy"
    Exit Sub
End If
Mth_XRmv A
If Trc Then
Debug.Print "Mth_XMov: Done.......................//"
End If
End Sub

Private Sub Z_Mth_XMov()
'Mth_XMov Md("Mth_"), "XX", Md("A_")
End Sub

Sub Md_XXMov_MthNy(A As CodeModule, MthNy$(), ToMd As CodeModule)
If Sz(MthNy) = 0 Then Exit Sub
Dim N
For Each N In MthNy
    MdMthNm_XMov A, CStr(N), ToMd
Next
End Sub




Private Sub Z()
Z_MdMthNm_XMov
Z_Mth_XMov
MIde_Mth_Op_Mov:
End Sub
