Attribute VB_Name = "MIde_Z_Pj_Dte"
Option Compare Binary
Option Explicit
Function FbPjDte(A) As Date
Static Y As New Access.Application
Y.OpenCurrentDatabase A
Y.Visible = False
FbPjDte = Acs_PjDte(Y)
Y.CloseCurrentDatabase
End Function

Function Pj_FfnPjDte(Pj_Ffn) As Date
Select Case True
Case IsFxa(Pj_Ffn): Pj_FfnPjDte = FileDateTime(Pj_Ffn)
Case IsFb(Pj_Ffn): Pj_FfnPjDte = FbPjDte(Pj_Ffn)
Case Else: Stop
End Select
End Function

Function Acs_PjDte(A As Access.Application)
Dim O As Date
Dim M As Date
M = Itr_MaxPrp(A.CurrentProject.AllForms, "DateModified")
O = Max(O, M)
O = Max(O, Itr_MaxPrp(A.CurrentProject.AllModules, "DateModified"))
O = Max(O, Itr_MaxPrp(A.CurrentProject.AllReports, "DateModified"))
Acs_PjDte = O
End Function
