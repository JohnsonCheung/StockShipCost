Attribute VB_Name = "MDao_Z_Prp_Tbl"
Option Compare Binary
Option Explicit
Private Sub ZZ_Dbt_Prp()
Tbl_XDrp "Tmp"
DoCmd.RunSQL "Create Table Tmp (F1 Text)"
Dbt_Prp(CurDb, "Tmp", "XX") = "AFdf"
Debug.Assert Dbt_Prp(CurDb, "Tmp", "XX") = "AFdf"
End Sub

Function Dbt_XCrt_Prp(A As Database, T, P$, V) As DAO.Property
Set Dbt_XCrt_Prp = A.TableDefs(T).CreateProperty(P, Val_DaoTy(V), V) ' will break if V=""
End Function

Function Dbt_Exist_Prp(A As Database, T, P$) As Boolean
Dbt_Exist_Prp = Itr_XHas_Nm(A.TableDefs(T).Properties, P)
End Function

Property Get Dbt_Prp(A As Database, T, P$)
If Not Dbt_Exist_Prp(A, T, P) Then Exit Property
Dbt_Prp = A.TableDefs(T).Properties(P).Value
End Property

Property Let Dbt_Prp(A As Database, T, P$, V)
If V = "" Then
    Dbt_PrpDrp A, T, P
    Exit Property
End If
If Dbt_Exist_Prp(A, T, P) Then
    A.TableDefs(T).Properties(P).Value = V
Else
    A.TableDefs(T).Properties.Append Dbt_XCrt_Prp(A, T, P, V)
End If
End Property

Sub Dbt_PrpDrp(A As Database, T, P$)
Dim I As DAO.Property, Prps As DAO.Properties
Set Prps = A.TableDefs(T).Properties
For Each I In Prps
    If I.Name = P Then
        Prps.Delete P
        Exit Sub
    End If
Next
End Sub
