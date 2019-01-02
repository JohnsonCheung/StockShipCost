Attribute VB_Name = "MIde_Dcl_Analysis"
Option Explicit
Option Compare Text

Function DclLin_T1$(A)
Dim L$: L = LTrim(A)
If L = "" Then Exit Function
If XTak_FstChr(L) = "'" Then Exit Function
DclLin_T1 = Lin_T1(A)
End Function
Function DclLy_T1ASet(A$()) As ASet
Dim L, O As ASet
Set O = EmpASet
For Each L In AyNz(A)
    ASet_XPush O, DclLin_T1(L)
Next
End Function
Function Md_DclLinT1Ay(A As CodeModule) As String()
Md_DclLinT1Ay = DclLy_T1ASet(Md_DclLy(A))
End Function
