Attribute VB_Name = "MIde_Z_Pj_Cmp_Dlt"
Option Compare Binary
Option Explicit
Sub Pj_XDlt_Md(A As VBProject, MdNm$)
If Not Pj_XHas_MdNm(A, MdNm) Then Exit Sub
A.VBComponents.Remove A.VBComponents(MdNm)
End Sub

Sub MdDlt(A As CodeModule)
Dim M$, P$, Pj As VBProject
    M = Md_Nm(A)
    Set Pj = Md_Pj(A)
    P = Pj.Name
Debug.Print QQ_Fmt("MdDlt: Before Md(?) is deleted from Pj(?)", M, P)
A.Parent.Collection.Remove A.Parent
Debug.Print QQ_Fmt("MdDlt: After Md(?) is deleted from Pj(?)", M, P)
End Sub
Sub MdRmv(A As CodeModule)
Dim C As VBComponent: Set C = A.Parent
C.Collection.Remove C
End Sub

Sub PjRmvMdNmPfx(A As VBProject, Pfx$)
Dim I
For Each I In Pj_MdAy(A, WhMd(Nm:=WhNm("^" & Pfx)))
    Md_XRen_RmvNmPfx CvMd(I), Pfx
Next
End Sub

Sub PjRmvMdPfx(A As VBProject, B As WhMd, MdPfx$)
Dim Md As CodeModule, M
For Each M In AyNz(Pj_MdAy(A, B))
    Set Md = M
    Md.Parent.Name = XRmv_Pfx(Md_Nm(A), MdPfx)
Next
End Sub

