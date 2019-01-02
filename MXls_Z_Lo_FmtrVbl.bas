Attribute VB_Name = "MXls_Z_Lo_FmtrVbl"
Option Compare Binary
Option Explicit

Function Qt_PrpLoFmtrVbl$(A As QueryTable)
Qt_PrpLoFmtrVbl = FbtStr_PrpLoFmtrVbl(Qt_FbtStr(A))
End Function

Function Qt_PrpLoFmtVbl$(A As QueryTable)
If IsNothing(A) Then Exit Function
Qt_PrpLoFmtVbl = FbtStr_PrpLoFmtrVbl(Qt_FbtStr(A))
End Function

Property Get Dbt_PrpLoFmtrVbl$(A As Database, T)
Dbt_PrpLoFmtrVbl = Dbt_Prp(A, T, "LoFmtrVbl")
End Property

Property Let Dbt_PrpLoFmtrVbl(A As Database, T, LoFmtrVbl$)
Dbt_Prp(A, T, "LoFmtrVbl") = LoFmtrVbl
End Property

Property Get Tbl_PrpLoFmtrVbl$(T)
Tbl_PrpLoFmtrVbl = Dbt_PrpLoFmtrVbl(CurDb, T)
End Property

Property Let Tbl_PrpLoFmtrVbl(T, LoFmtrVbl$)
Dbt_PrpLoFmtrVbl(CurDb, T) = LoFmtrVbl
End Property

Function Lo_PrpLoFmtlVbl$(A As ListObject)
Lo_PrpLoFmtlVbl = Qt_PrpLoFmtrVbl(Lo_Qt(A))
End Function

Property Get Fbt_PrpLoFmtrVbl$(A$, T$)
Fbt_PrpLoFmtrVbl = Dbt_PrpLoFmtrVbl(Fb_Db(A), T)
End Property

Property Let Fbt_PrpLoFmtrVbl(A$, T$, LoFmtrVbl$)
Dbt_PrpLoFmtrVbl(Fb_Db(A), T) = LoFmtrVbl
End Property

Function FbtStr_PrpLoFmtrVbl$(FbtStr$)
Dim Fb$, T$
FbtStr_XAsg FbtStr, Fb, T
FbtStr_PrpLoFmtrVbl = Fbt_PrpLoFmtrVbl(Fb, T)
End Function

Property Get TpMainPrpLoFmtrVbl$()
'TpMainPrpLoFmtrVbl = Lo_PrpLoFmtlVbl(TpMainLo)
End Property
