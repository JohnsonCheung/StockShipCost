Attribute VB_Name = "MDao_Z_Db_Dbt_Is"
Option Compare Binary
Option Explicit

Function Dbt_IsLnk(A As Database, T) As Boolean
Dbt_IsLnk = Dbt_IsFbLnk(A, T) Or Dbt_IsFxLnk(A, T)
End Function

Function Dbt_IsFbLnk(A As Database, T) As Boolean
Dbt_IsFbLnk = XHas_Pfx(Dbt_CnStr(A, T), ";Database=")
End Function

Function Dbt_IsFxLnk(A As Database, T) As Boolean
Dbt_IsFxLnk = XHas_Pfx(Dbt_CnStr(A, T), "Excel")
End Function

Function Dbt_IsSys(A As Database, T) As Boolean
Dbt_IsSys = A.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbSystemObject
End Function

Function Dbt_IsXls(A As Database, T) As Boolean
Dbt_IsXls = XHas_Pfx(Dbt_CnStr(A, T), "Excel")
End Function


