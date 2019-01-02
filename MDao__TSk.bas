Attribute VB_Name = "MDao__TSk"
Option Compare Binary
Option Explicit

Function Tbl_XHas_SkValAp(T, ParamArray SkValAp()) As Boolean
Dim Sk(): Sk = SkValAp
Tbl_XHas_SkValAp = Dbt_XHas_SkVV(CurDb, T, Sk)
End Function

Sub TblValAp_XIns(T, ParamArray ValAp())
Dim Vy(): Vy = ValAp
DbtVy_XIns CurDb, T, Vy
End Sub

Function Dbt_XHas_SkVV(A As Database, T, SkAv()) As Boolean
Dbt_XHas_SkVV = Dbq_XHas_Rec(A, QSel_Fm(T, Dbt_SkAvWhStr(A, T, SkAv)))
End Function

Function TF_Val_BySkVal(T, F, SkVal) ' S is Ssk-Value  (Ssk is single-field-secondary-key)
TF_Val_BySkVal = Dbtf_Val_BySkVal(CurDb, T, F, SkVal)
End Function

Function Dbtf_Val_BySkVal(A As Database, T, F, SkVal) ' S is Ssk-Value (Ssk is single-field-secondary-key)
Dim B$
B = DbtSkVV_BExpr(A, T, SkVal)
Dbtf_Val_BySkVal = Dbq_Val(A, QSel_FF_Fm_Wh(F, T, B))
End Function
