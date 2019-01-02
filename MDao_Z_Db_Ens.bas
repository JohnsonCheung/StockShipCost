Attribute VB_Name = "MDao_Z_Db_Ens"
Option Compare Binary
Option Explicit

Sub Td_XEns(A As DAO.TableDef)
DbTd_XEns CurDb, A
End Sub

Sub DbSchm_XEns(A As DAO.TableDef, SchmLines$)
DbTd_XEns CurDb, A
End Sub

Sub DbTd_XEns(A As Database, B As DAO.TableDef)
If Tbl_Exist(B.Name) Then
    Td_IsEq_XAss CurrentDb.TableDefs(B.Name), B
Else
    CurrentDb.TableDefs.Append B
End If
End Sub



