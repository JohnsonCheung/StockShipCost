Attribute VB_Name = "MDao_TF_Id"
Option Compare Binary
Option Explicit

Property Get Dbtfi_Val(A As Database, T, F, Id&)
Dbtfi_Val = Dbtfi_Rs(A, T, Id, F).Fields(0).Value
End Property

Function Dbtfi_Rs(A As Database, T, I&, F) As DAO.Recordset
Q = QQ_Fmt("Select [?] From [?] where [?Id]=?", F, T, T, I)
Set Dbtfi_Rs = A.OpenRecordset(Q)
End Function

Property Let Dbtfi_Val(A As Database, T, F, Id&, V)
With Dbtfi_Rs(A, T, Id, F)
    .Edit
    .Fields(0).Value = V
    .Update
End With
End Property

Property Let Tfi_Val(T, F, Id&, V)
Dbtfi_Val(CurDb, T, F, Id) = V
End Property

Property Get Tfi_Val(T, F, Id&)
Tfi_Val = Dbtfi_Val(CurDb, T, F, Id)
End Property

