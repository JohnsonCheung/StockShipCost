Attribute VB_Name = "MDao_Z_Db_DbInf_Stru"
Option Compare Binary
Option Explicit

Function DbInfDtStru(A As Database) As Dt
Dim T$, TT, Dry(), Des$, NRec&, Stru$
For Each TT In Db_Tny(A)
    T = TT
    Des = Dbt_Des(A, T)
    Stru = RmvT1(Dbt_Stru(A, T))
    NRec = Dbt_NRec(A, T)
    PushI Dry, Array(T, NRec, Des, Stru)
Next
Set DbInfDtStru = New_Dt("Tbl", "Tbl NRec Des", Dry)
End Function

Function Db_Stru$(A As Database)
Db_Stru = Dbtt_Stru(A, Db_Tny(A))
End Function

Sub Db_Stru_Dmp(A As Database)
D Db_Stru(A)
End Sub

Property Get Stru$()
Stru = CurDb_Stru
End Property

Function StruFld(ParamArray Ap()) As Drs
Dim Dry(), Av(), Ele$, LikFF, LikFld, X
Av = Ap
For Each X In Av
    Lin_TRstAsg X, Ele, LikFF
    For Each LikFld In Ssl_Sy(LikFF)
        PushI Dry, Array(Ele, LikFld)
    Next
Next
Set StruFld = New_Drs("Ele FldLik", Dry)
End Function

Function TT_Stru$(TT)
TT_Stru = Dbtt_Stru(CurDb, TT)
End Function

Sub TTStruDmp(TT)
D TT_Stru(TT)
End Sub
