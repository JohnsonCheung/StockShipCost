Attribute VB_Name = "MIde_Z_Pj_Activate"
Option Compare Binary
Option Explicit

Sub Pj_XAct(A As VBProject)
Set Pj_Vbe(A).ActiveVBProject = A
End Sub

Sub PjNm_XAct(PjNm$)
Set CurVbe.ActiveVBProject = Pj(PjNm)
End Sub
