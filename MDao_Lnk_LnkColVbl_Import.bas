Attribute VB_Name = "MDao_Lnk_LnkColVbl_Import"
Option Compare Binary
Option Explicit

Sub Dbt_XImp_ByLnkVblCol(A As Database, T, LnkColVbl$, Optional WhBExpr$)
'Create Tbl-[#I*] T from Tbl[>*]
If XTak_FstChr(T) <> ">" Then
    Debug.Print "T must have first char = '>'"
    Stop
End If
Dim Into$
    Into = "#I" & XRmv_FstChr(T)
Dbt_XDrp A, Into
Q = LnkColVbl_ImpSql(LnkColVbl, T, Into, WhBExpr): A.Execute Q
Debug.Print Q
Debug.Print
End Sub

Private Function LnkColVbl_ImpSql$(A$, T, Into, Optional WhBExpr$)
Dim Ny$(), ExtNy$()
    LnkColVbl_NyExtNy_XAsg A, Ny, ExtNy
LnkColVbl_ImpSql = QSel_FF_ExprAy_Into_Fm_OWh(Ny, ExtNy, Into, T, WhBExpr)
End Function


