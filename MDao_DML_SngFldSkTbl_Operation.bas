Attribute VB_Name = "MDao_DML_SngFldSkTbl_Operation"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_DML_SngFldSkTbl_Operation."

Private Sub Z_Dbt_IntAy()
Dim Db As Database
Set Db = TmpDb
Db.Execute "Create Table [Tbl-AA] ([Fld-AA] Int)"
Db.Execute TblSkFF_CrtSkSql("Tbl-AA", "Fld-AA")
Db.Execute "Insert Into [Tbl-AA] ([Fld-AA]) Values(1)"
Db.Execute "Insert Into [Tbl-AA] ([Fld-AA]) Values(3)"
Db.Execute "Insert Into [Tbl-AA] ([Fld-AA]) Values(5)"
SskVal_XIns_Dbt New_ASet_AP(1, 2, 3, 4, 5, 6, 7), Db, "Tbl-AA"
Act = Dbtf_IntAy(Db, "Tbl-AA", "Fld-AA")
Ept = Ap_IntAy(1, 2, 3, 4, 5, 6, 7)
C
End Sub

Function Dbt_SskFldNm$(A As Database, T)
Const CSub$ = CMod & "DbtSngSkFld"
Dim Sk$(): Sk = Dbt_SkFny(A, T)
If Sz(Sk) = 1 Then Dbt_SskFldNm = Sk(0): Exit Function
XThw CSub, "Given Sk Sz <>1", "Db T Sk SkSz", Db_Nm(A), T, Sk, Sz(Sk)
End Function

Sub SskVal_XDlt_Dbt(SskVal As ASet, Db As Database, T) _
'Delete Db-T record for those record's Sk not in Ay-A, _
'Assume T has single-fld-sk
Dim Q$, Excess As ASet
Dim SskFldNm$
    SskFldNm = Dbt_SskFldNm(Db, T)
'Set Excess = ASet_XMinus(Dbt_SkFny  SkASet(Db, T), SkASet)
If ASet_IsEmp(SskVal) Then XThw CSub, "Given SskVal cannot be empty", "Db T", Db_Nm(Db), T
DbSqy_XRun Db, QDlt_Fm_F_InASet(T, SskFldNm, Excess)
End Sub

Function Dbt_SskVal(A As Database, T) As ASet
Set Dbt_SskVal = Dbtf_ASet(A, T, Dbt_SskFldNm(A, T))
End Function

Function Dbt_XMinus_SskVal(A As Database, T, SskVal As ASet) As ASet
Set Dbt_XMinus_SskVal = ASet_XMinus(Dbt_SskVal(A, T), SskVal)
End Function

Sub SskVal_XIns_Dbt(A As ASet, Db As Database, T$) _
'Insert Ay-A into Db-T _
'Assume T has single-fld-sk and can be inserted by just giving such Sk-value
Const CSub$ = CMod & "AyInsDbt"
Dim SkFld$
    Dim Sk$(): Sk = Dbt_SkFny(Db, T)
    If Sz(Sk) <> 1 Then XThw CSub, "Dbt does not have Sk of sz=1", "Db T Sk-Sz Sk", Db_Nm(Db), T, Sz(Sk), Sk
    SkFld = Sk(0)

ZCrtTmpTbl:
    Dim TmpTbl$
    TmpTbl = TmpNm
    DbtFF_XCrt_TmpTbl Db, T, SkFld, TmpTbl

ZInsAyToTmpTbl:
    Dim X
    With Dbt_Rs(Db, TmpTbl)
        For Each X In AyNz(A)
            .AddNew
            .Fields(0).Value = X
            .Update
        Next
    End With

ZInsIntoTarTbl:
    Q = QQ_Fmt("Insert Into [?] select x.[?] from [?] x left join [?] a on x.[?] = a.[?] where a.[?] is null", _
        T, SkFld, TmpTbl, T, SkFld, SkFld, SkFld)
    Db.Execute Q
    
ZDltTmpTbl:
    Dbt_XDrp Db, TmpTbl
End Sub




Private Sub Z()
MDao_DML_SngFldSkTbl_Operation:
End Sub
