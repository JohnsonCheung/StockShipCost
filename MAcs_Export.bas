Attribute VB_Name = "MAcs_Export"
Option Compare Binary
Option Explicit

Sub Acs_XExp_Frm(A As Access.Application)
Dim Nm$, P$, I
P = Acs_SrcPth(A)
For Each I In Acs_FrmNy(A)
    SaveAsText acForm, Nm, P & Nm & ".Frm.Txt"
Next
End Sub

Function Acs_FrmNy(A As Access.Application) As String()
Acs_FrmNy = Itr_Ny(A.CodeProject.AllForms)
End Function

Function Acs_SrcPth$(A As Access.Application)
Acs_SrcPth = Pj_SrcPth(A.Vbe.ActiveVBProject)
End Function
