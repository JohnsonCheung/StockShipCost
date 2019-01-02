Attribute VB_Name = "MDao_Lnk_Tbl_Fb"
Option Compare Binary
Option Explicit
Sub Db_XLnk_Fb(A As Database, Fb$, Tny0, Optional SrcTny0)
Dim Tny$(): Tny = CvNy(Tny0)              ' Src_Tny
Dim Src$(): Src = CvNy(Dft(SrcTny0, Tny0)) ' Tar_Tny
Ass Sz(Tny) > 0
Ass Sz(Src) = Sz(Tny)
Dim J%
For J = 0 To UB(Tny)
    Dbt_XLnk_Fb A, Tny(J), Fb$, Src(J)
Next
End Sub

Function Db_XLnk_Tny(A As Database) As String()
Db_XLnk_Tny = Itr_PrpSy_WhPrepPrpTrue(A.TableDefs, "Any_CnStr_Td", "Name")
End Function

Function Dbt_XLnk_Fb(A As Database, T, Fb$, Optional Fbt0$) As String()
Dim Fbt$, Cn$
Cn = ";Database=" & Fb
Fbt = DftStr(Fbt0, T)
Dbt_XLnk_Fb = Dbt_XLnk(A, T, Fbt, Cn)
End Function
