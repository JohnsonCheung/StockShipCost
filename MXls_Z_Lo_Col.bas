Attribute VB_Name = "MXls_Z_Lo_Col"
Option Compare Binary
Option Explicit
Sub Lc_XSet_Fml(A As ListObject, C, Fml$)
Lc_Rg(A, C).Formula = Fml
End Sub

Sub Lc_XSet_Cor(A As ListObject, C, Cor&)
Lc_Rg(A, C).Interior.Color = Cor
End Sub

Sub Lc_XSet_Fmt(A As ListObject, C, Fmt$)
Lc_Rg(A, C).Formula = Fmt
End Sub

Sub Lc_XSet_Lvl(A As ListObject, C, Optional Lvl As Byte = 2)
Lc_Rg(A, C).EntireColumn.OutlineLevel = Lvl
End Sub

Sub Lc_XSet_Wdt(A As ListObject, C, W%)
Lc_Rg(A, C).EntireCumn.ColumnWidth = W
End Sub

Sub Lc_XSet_Align(A As ListObject, C, B As XlHAlign)
Lc_Rg(A, C).HorizontalAlignment = B
End Sub

Private Function Lc_Rg(A As ListObject, C) As Range
Set Lc_Rg = A.ListColumns(C).DataBodyRange
End Function

Sub Lc_XSet_BdrLeft(A As ListObject, C)
Rg_XBdr_L Lc_Rg(A, C)
End Sub

Sub Lc_XSet_BdrRight(A As ListObject, C)
Rg_XBdr_L Lc_Rg(A, C)
End Sub

Sub Lc_XSet_Tot(A As ListObject, C, B As XlTotalsCalculation)
A.ListColumns(C).TotalsCalculation = B
End Sub
