Attribute VB_Name = "MIde_Z_Md_Op_Ren"
Option Compare Binary
Option Explicit
Sub XRen_Md(NewNm$)
CurMd.Name = NewNm
End Sub

Sub Md_XRen(A As CodeModule, NewNm$)
Dim Nm$: Nm = Md_Nm(A)
If NewNm = Nm Then
    Debug.Print QQ_Fmt("Md_XRen: Given Md-[?] name and NewNm-[?] is same", Nm, NewNm)
    Exit Sub
End If
If Pj_XHas_MdNm(Md_Pj(A), NewNm) Then
    Debug.Print QQ_Fmt("Md_XRen: Md-[?] already exist.  Cannot rename from [?]", NewNm, Md_Nm(A))
    Exit Sub
End If
Md_Cmp(A).Name = NewNm
Debug.Print QQ_Fmt("Md_XRen: Md-[?] renamed to [?] <==========================", Nm, NewNm)
End Sub

Private Sub Z_Md_XRen()
Md_XRen Md("A_Rs1"), "A_Rs"
End Sub

Sub Pj_XRen_Md_ByPfx(A As VBProject, FmMdPfx$, ToMdPfx$)
Dim CvNy$()
Dim Ny$()
'    Ny = Pj_MdNy(A, "^" & FmMdPfx)
    CvNy = AyMapAsgSy(Ny, "RplPfx", FmMdPfx, ToMdPfx)
Dim MdAy() As CodeModule
    Dim MdNm
    Dim Md As CodeModule
    For Each MdNm In Ny
        Set Md = Pj_Md(A, CStr(MdNm))
        PushObj MdAy, Md
    Next
Dim I%, U%
    For I = 0 To UB(CvNy)
        Md_XRen MdAy(I), CvNy(I)
    Next
End Sub

Private Sub Z_Pj_XRen_Md_ByPfx()
Pj_XRen_Md_ByPfx CurPj, "A_", ""
End Sub

Private Sub Z()
Z_Md_XRen
Z_Pj_XRen_Md_ByPfx
MIde_Z_Md_Op_Ren:
End Sub
