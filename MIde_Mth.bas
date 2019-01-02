Attribute VB_Name = "MIde_Mth"
Option Compare Binary
Option Explicit

Function Md_XEns_Mth(A As CodeModule, MthNm$, NewMthLines$)
Dim OldMthLines$: OldMthLines = MdMthNm_MthLines(A, MthNm)
If OldMthLines = NewMthLines Then
    Debug.Print QQ_Fmt("Md_XEns_Mth: Mth(?) in Md(?) is same", MthNm, Md_Nm(A))
End If
MdMthNm_XRmv A, MthNm
Md_XApp_Lines A, NewMthLines
Debug.Print QQ_Fmt("Md_XEns_Mth: Mth(?) in Md(?) is replaced <=========", MthNm, Md_Nm(A))
End Function


Function Md_MthPfxAy(A As CodeModule) As String()
Dim N
For Each N In AyNz(Md_MthNy(A))
    PushNoDup Md_MthPfxAy, MthPfx(N)
Next
End Function

Function Md_XHas_MthNm(A As CodeModule, MthNm$, Optional WhMdy$, Optional WhKd$) As Boolean
Md_XHas_MthNm = Src_XHas_MthNm(Md_BdyLy(A), MthNm, WhMdy, WhKd)
End Function

Function MdHasTstSub(A As CodeModule) As Boolean
Dim I
For Each I In Md_Ly(A)
    If I = "Friend Sub Z()" Then MdHasTstSub = True: Exit Function
    If I = "Sub Z()" Then MdHasTstSub = True: Exit Function
Next
End Function



Function MdMdMthNm_FmCntAy(A As CodeModule, MthNm$) As FmCnt()
MdMdMthNm_FmCntAy = SrcMthNm_FmCntAy(Md_Src(A), MthNm)
End Function

Private Sub Z_MdMdMthNm_FmCntAy()
Dim A() As FmCnt: A = MdMdMthNm_FmCntAy(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FmCntDmp A(J)
Next
End Sub

Function MdMthAy(A As CodeModule) As Mth()
Dim N
For Each N In AyNz(Md_MthNy(A))
    PushObj MdMthAy, New_Mth(A, N)
Next
End Function



Function MdMthKeyLinesDic1(A As CodeModule) As Dictionary
'To be delete
'Dim Pfx$: Pfx = Mth_PjNm(A) & "." & Md_Nm(A) & "."
'Set MdMthKeyLinesDic = DicAddKeyPfx(SrcMthKeyLinesDic(Md_Src(A)), Pfx)
End Function

Function MdMthKy(A As CodeModule) As String()
MdMthKy = Ay_XAdd_Pfx(Src_MthKy(Md_Src(A)), Md_DNm(A) & ".")
End Function





Private Sub Z()
Z_MdMdMthNm_FmCntAy
MIde_Mth:
End Sub
