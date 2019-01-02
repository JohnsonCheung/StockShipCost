Attribute VB_Name = "MDao_Z_Db_Dbq"
Option Compare Binary
Option Explicit
Private Sub Z_Dbq_Val()
Ept = CByte(18)
Act = Sql_Val("Select Y from [^YM]")
C
End Sub

Function Dbq_Val(A As Database, Q)
Dbq_Val = Rs_Val(Dbq_Rs(A, Q))
End Function

Function Dbt_FstSndColDic(A As Database, T) As Dictionary
Dim F$, S$
Dbt_FstSndColNm A, T, F, S
Set Dbt_FstSndColDic = Dbq_Dic(A, QQ_Fmt("Select [?],[?] from [?]", F, S, T))
End Function
Sub Dbt_FstSndColNm(A As Database, T, OFstFldNm$, OSndFldNm$)
With A.TableDefs(T)
    OFstFldNm = .Fields(0).Name
    OSndFldNm = .Fields(1).Name
End With
End Sub

Function Dbq_XHas_Rec(A As Database, Sql) As Boolean
Dbq_XHas_Rec = Rs_IsNoRec(Dbq_Rs(A, Sql))
End Function

Sub Dbq_XBrw(A As Database, Sql$)
Drs_XBrw Dbq_Drs(A, Sql)
End Sub

Function Dbq_Dr(A As Database, Q$) As Variant()
Dbq_Dr = Rs_Dr(A.OpenRecordset(Q))
End Function

Function QQ_Drs(QQ, ParamArray Ap()) As Drs
Dim Av(): Av = Ap
Set QQ_Drs = Dbq_Drs(CurDb, QQ_FmtAv(QQ, Av))
End Function

Function Dbq_Drs(A As Database, Q) As Drs
Set Dbq_Drs = Rs_Drs(Dbq_Rs(A, Q))
End Function

Function Dbq_Dry(A As Database, Q) As Variant()
Dbq_Dry = Rs_Dry(Dbq_Rs(A, Q))
End Function

Function DbqDTim$(A As Database, Sql)
DbqDTim = Dte_DTim(Dbq_Val(A, Sql))
End Function

Function Dbq_IntAy(A As Database, Q) As Integer()
Dbq_IntAy = Rs_IntAy(Dbq_Rs(A, Q))
End Function

Function Dbq_Lng&(A As Database, Sql)
Dbq_Lng = Dbq_Val(A, Sql)
End Function

Function Dbq_LngAy(A As Database, Sql) As Long()
Dbq_LngAy = Rs_LngAy(A.OpenRecordset(Sql))
End Function

Function Dbq_Rs(A As Database, Q) As DAO.Recordset
Set Dbq_Rs = A.OpenRecordset(Q)
End Function

Sub Dbq_XRun(A As Database, Q)
A.Execute Q
End Sub

Sub Db_XRun(A As Database, Sql_or_Sqy)
If Not IsArray(Sql_or_Sqy) Then
    A.Execute Sql_or_Sqy
    Exit Sub
End If
Dim Q
For Each Q In Sql_or_Sqy
    A.Execute Q
Next
End Sub

Function Dbq_Sy(A As Database, Q) As String()
Dbq_Sy = Rs_Sy(A.OpenRecordset(Q))
End Function

Function Dbq_Tim(A As Database, Q) As Date
Dbq_Tim = Dbq_Val(A, Q)
End Function

Private Sub Z()
Z_Dbq_Val
MDao_Z_Db_Dbq:
End Sub
