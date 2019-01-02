Attribute VB_Name = "MAcs_Acs_Tbl"
Option Compare Binary
Option Explicit

Sub AcsT_Cls(A As Access.Application, T)
A.DoCmd.Close acTable, T, acSaveYes
End Sub

Sub AcsTT_Cls(A As Access.Application, TT)
Ay_XDoPX CvNy(TT), "AcstCls", A
End Sub

Sub Acs_XCls_AllTbl(A As Access.Application)
Dim T As AccessObject
For Each T In A.CodeData.AllTables
    A.DoCmd.Close acTable, T.Name
Next
End Sub
