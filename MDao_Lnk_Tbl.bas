Attribute VB_Name = "MDao_Lnk_Tbl"
Option Compare Binary
Option Explicit

Function Dbt_XLnk(A As Database, T, S$, Cn$) As String()
On Error GoTo X
Dim Td As New DAO.TableDef
Dbtt_XDrp A, T
With Td
    .Connect = Cn
    .Name = T
    .SourceTableName = S
    A.TableDefs.Append Td
End With
Exit Function
X:
Dim E$
E = Err.Description
Dbt_XLnk = FunMsgNyAp_Ly(CSub, "Cannot link", "Db Tbl SrcTbl CnStr Er", Db_Nm(A), T, S, Cn, E)
End Function

Function Dbt_LnkVbl$(A As Database, T)
Dim O$
O = Dbt_LnkVbl_FXW(A, T): If O <> "" Then Dbt_LnkVbl = "LnkFx|" & O: Exit Function
O = Dbt_LnkVbl_FBT(A, T): If O <> "" Then Dbt_LnkVbl = "LnkFb|" & O: Exit Function
Dbt_LnkVbl = "Lcl|" & A.Name & "|" & T
End Function

Function Dbt_LnkVbl_FXW$(A As Database, T)
If Dbt_IsFxLnk(A, T) Then Dbt_LnkVbl_FXW = Dbt_LnkVbl_RAW(A, T)
End Function

Function Dbt_LnkVbl_FBT$(A As Database, T)
If Dbt_IsFbLnk(A, T) Then
    Dbt_LnkVbl_FBT = Dbt_LnkVbl_RAW(A, T)
End If
End Function

Function Dbt_LnkVbl_RAW$(A As Database, T)
Dim Cn$, X$, Y$, Y1$
Cn = Dbt_CnStr(A, T)
X = XTak_BefOrAll(XTak_Aft(Cn, "DATABASE="), ";")
Y = A.TableDefs(T).SourceTableName
Y1 = RmvSfx(Y, "$")
Dbt_LnkVbl_RAW = X & "|" & Y1
End Function
