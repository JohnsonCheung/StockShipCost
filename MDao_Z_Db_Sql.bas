Attribute VB_Name = "MDao_Z_Db_Sql"
Option Compare Binary
Option Explicit
Function Any_Sql(A) As Boolean
Any_Sql = Dbq_XHas_Rec(CurDb, A)
End Function

Function Sql_Dry(A) As Variant()
Sql_Dry = Dbq_Dry(CurDb, A)
End Function

Function Sql_Fny(A) As String()
Sql_Fny = Rs_Fny(Sql_Rs(A))
End Function

Function Sql_Lng(A)
Sql_Lng = Dbq_Lng(CurDb, A)
End Function

Function Sql_LngAy(A) As Long()
Sql_LngAy = Dbq_LngAy(CurDb, A)
End Function


Function Sql_Rs(A) As DAO.Recordset
Set Sql_Rs = CurDb.OpenRecordset(A)
End Function

Sub Sql_XRun(A)
CurDb.Execute A
End Sub

Function Sql_Sy(A) As String()
Sql_Sy = Dbq_Sy(CurDb, A)
End Function

Function Sql_Val(A)
Sql_Val = Dbq_Val(CurDb, A)
End Function

Private Sub ZZ_Sql_Fny()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
Ay_XDmp Sql_Fny(S)
End Sub

Private Sub ZZ_Sql_Rs()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
Ay_XBrw Rs_CsvLy(Sql_Rs(S))
End Sub

Private Sub Z_Sql_Sy()
Ay_XDmp Sql_Sy("Select Distinct UOR from [>Imp]")
End Sub

