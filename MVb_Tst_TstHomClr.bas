Attribute VB_Name = "MVb_Tst_TstHomClr"
Option Compare Binary
Option Explicit

Sub TstHomClr() ' Rmv-Empty-Pth Rmk-Pth-As-At
Pth_XRmv_EmpSubDirR TstHom
Ren_Pj_Pth_AsAt
Ren_MdPth_AsAt
Ren_MthPth_AsAt
Ren_CasPth_AsAt
End Sub
Private Sub Ren_Pj_Pth_AsAt()

End Sub
Private Sub Ren_MdPth_AsAt()

End Sub
Private Sub Ren_MthPth_AsAt()

End Sub
Private Sub Ren_CasPth_AsAt()
Ren CasPthAy
End Sub
Private Property Get CasPthAy() As String()

End Property
Private Sub Ren(PthAy)
Dim I
For Each I In AyNz(PthAy)
    PthRenXAdd_Pfx I, "@"
Next
End Sub
