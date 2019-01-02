Attribute VB_Name = "MDao_Z_Db_Dbt_Sk"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Z_Db_Dbt_Sk."

Sub DbtAssSk(A As Database, T)
Const CSub$ = CMod & "DbtAssSk"
Dim SkIdx As DAO.Index, I As DAO.Index
Set SkIdx = Itr_FstItmNm(A.TableDefs(T).Indexes, "SecondaryKey")
Select Case True
Case Not IsNothing(SkIdx)
    If Not SkIdx.Unique Then
        XThw CSub, "There is SecondaryKey idx, but it is not uniq", "Db Tbl Key", Db_Nm(A), T, Idx_Fny(SkIdx)
    End If
Case Else
    Set I = Dbt_FstUniqIdx(A, T)
    If Not IsNothing(I) Then
        XThw CSub, "No SecondaryKey, but there is uniq idx, it should name as SecondaryKey", _
        "Db Tbl UniqKeyNm UniqKeyFld", _
        Db_Nm(A), T, I.Name, Idx_Fny(I)
    End If
End Select
End Sub

Function Dbt_FstUniqIdx(A As Database, T) As DAO.Index
Set Dbt_FstUniqIdx = Itr_FstItm_PrpTrue(A.TableDefs(T).Indexes, "Unique")
End Function

Private Sub ZZ()
Dim A As Database
Dim B
Dim XX
DbtAssSk A, B
Dbt_FstUniqIdx A, B
End Sub

Private Sub Z()
End Sub
