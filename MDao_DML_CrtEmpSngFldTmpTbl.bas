Attribute VB_Name = "MDao_DML_CrtEmpSngFldTmpTbl"
Option Compare Binary
Option Explicit

Function DbtFF_XCrt_TmpTbl$(A As Database, T, FF, TmpTbl$)
A.Execute QSel_FF_Into_Fm_OWh(FF, TmpTbl, T, "False")
End Function
