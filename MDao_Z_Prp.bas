Attribute VB_Name = "MDao_Z_Prp"
Option Compare Binary
Option Explicit
Function Prps_XHas_Prp(Prps As DAO.Properties, P) As Boolean
Prps_XHas_Prp = Itr_XHas_Nm(Prps, P)
End Function

Function PrpsP_Val(A As DAO.Properties, PrpNm$)
On Error Resume Next
PrpsP_Val = A(PrpNm).Value
End Function


Property Let Dbtf_Prp(A As Database, T, F, P, V)
'If IsEmpty(V) Then
'    If Dbtf_HasPrp(A, T, F, P) Then
'        A.TableDefs(T).Fields(T).Properties.Delete P
'    End If
'    Exit Function
'End If
If Dbtf_HasPrp(A, T, F, P) Then
    A.TableDefs(T).Fields(F).Properties(P).Value = V
Else
    With A.TableDefs(T)
        .Fields(F).Properties.Append .CreateProperty(P, Val_DaoTy(V), V)
    End With
End If
End Property

Property Get Dbtf_Prp(A As Database, T, F, P)
If Not Dbtf_HasPrp(A, T, F, P) Then Exit Property
Dbtf_Prp = A.TableDefs(T).Fields(F).Properties(P).Value
End Property

Function Dbtf_HasPrp(A As Database, T, F, P) As Boolean
Dbtf_HasPrp = Itr_XHas_Nm(A.TableDefs(T).Fields(F).Properties, P)
End Function

Property Let TblPrp(T, P$, V)
Dbt_Prp(CurDb, T, P) = V
End Property
