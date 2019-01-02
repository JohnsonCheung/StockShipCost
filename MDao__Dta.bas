Attribute VB_Name = "MDao__Dta"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao__Dta."
Function Ds_InsDbSqy(A As Ds, Db As Database) As String()
Dim Dt
For Each Dt In AyNz(A.DtAy)
   PushIAy Ds_InsDbSqy, Dt_InsDbSqy(CvDt(Dt), Db)
Next
End Function


Sub Ds_XIns_Db(A As Ds, Db As Database)
DbSqy_XRun Db, Ds_InsDbSqy(Db, A)
End Sub

Sub DtRplDb(A As Dt, Db As Database)
Dim T$: T = A.DtNm
Db.Execute QDlt_Fm_OWh(T)
DtInsDb A, Db
End Sub


Sub DsRplDb(A As Ds, Db As Database)
Dim T
For Each T In A.DtAy
    DtRplDb CvDt(T), Db
Next
End Sub

Sub Drs_XEns_Dbt(A As Drs, Db As Database, T$)
If Dbt_Exist(Db, T) Then Exit Sub
Db.Execute DrsCrtTblSql(A, T)
End Sub

Sub Drs_XEns_RplDbt(A As Drs, Db As Database, T$)
Drs_XEns_Dbt A, Db, T
Drs_XRpl_Dbt_Wh A, Db, T
End Sub

Sub DrsInsDbt(A As Drs, Db As Database, T$)
GoSub X
Dim Sqy$(): GoSub X_Sqy
DbSqy_XRun Db, Sqy
Exit Sub
X:
    Dim Dry, Fny$(), Sk$()
    Fny = A.Fny
    Dry = A.Dry
    Sk = Dbt_SkFny(Db, T)
    Return
X_Sqy:
    Dim Dr
    For Each Dr In AyNz(Dry)
        Push Sqy, InsSql(T, Fny, Dr)
    Next
    Return
End Sub

Sub Drs_XInsUpd_Dbt(A As Drs, Db As Database, T$)
GoSub X
Dim Ins As Drs, Upd As Drs: GoSub X_Ins_Upd
DrsInsDbt Ins, Db, T
Drs_XUpd_Dbt Upd, Db, T
Exit Sub
X:
    Dim Sk$()
    Dim SkIxAy&()
    Dim Fny$()
    Dim Dry()
    Sk = Dbt_SkFny(Db, T)
    Fny = A.Fny
    SkIxAy = Ay_IxAy(Fny, Sk)
    Dry = A.Dry
    Return
X_Ins_Upd:
    Dim IDry(), UDry(): GoSub X_IDry_UDry
    Set Ins = New_Drs(Fny, IDry)
    Set Upd = New_Drs(Fny, UDry)
    
    Return
X_IDry_UDry:
    Dim Dr, IsIns As Boolean, IsUpd As Boolean
    For Each Dr In Dry
        GoSub X_IsIns_IsUpd
        Select Case True
        Case IsIns: Push IDry, Dr
        Case IsUpd: Push UDry, Dr
        End Select
    Next
    Return
X_IsIns_IsUpd:
    IsIns = False
    IsUpd = False
    Dim SkVV(), Sql$, DbDr()
    SkVV = DrSel(CvAy(Dr), SkIxAy)
    Sql = QSel_FF_Fm_WhFny_EqVy(Fny, T, Sk, SkVV)
    DbDr = Dbq_Dr(Db, Sql)
    If Sz(DbDr) = 0 Then IsIns = True: Return
    If Not Ay_IsEq(DbDr, Dr) Then IsUpd = True: Return
    Return
End Sub

Private Sub Z_DtInsDbSqy()
'Tmp1Tbl_XEns
Stop
Dim Db As Database
Dim Dt As Dt: Dt = Dbt_Dt(CurDb, "Tmp1")
Dim O$(): O = Dt_InsDbSqy(Dt, Db)
Stop
End Sub

Sub DtInsDb(A As Dt, Db As Database)
If Sz(A.Dry) = 0 Then Exit Sub
Dim Rs As DAO.Recordset, Dr
Set Rs = Db.TableDefs(A.DtNm).OpenRecordset
For Each Dr In A.Dry
    Dr_XIns_Rs Dr, Rs
Next
Rs.Close
End Sub

Function Dt_InsDbSqy(A As Dt, Db As Database) As String()
Dim SimTyAy() As eSimTy
SimTyAy = Dbt_SimTyAy(Db, A.DtNm)
Dim ValTp$
   ValTp = SimTyAy_InsValTp(SimTyAy)
Dim Tp$
   Dim T$, F$
   T = A.DtNm
   F = JnComma(A.Fny)
   Tp = QQ_Fmt("Insert into [?] (?) values(?)", T, F, ValTp)
Dim O$()
   Dim Dr
   ReDim O(UB(A.Dry))
   Dim J%
   J = 0
   For Each Dr In A.Dry
       O(J) = QQ_FmtAv(Tp, Dr)
       J = J + 1
   Next
Dt_InsDbSqy = O
End Function
Sub Dbt_XClr(A As Database, T)
Dbq_XRun A, QDlt_Fm_OWh(T)
End Sub
Sub Drs_XRpl_Dbt(A As Drs, Db As Database, T)
Dbt_XClr A, T
Drs_XIns_Dbt A, Db, T
End Sub
Sub Drs_XIns_Dbt(A As Drs, Db As Database, T)

End Sub
Function Dbt_Drs_Wh(A As Database, T) As Drs

End Function
Function Drs_XMinus(A As Drs, B As Drs, FF) As Drs

End Function
Sub Drs_XIup_Dbt(A As Drs, Db As Database, T)
Dim Ins As Drs
Dim Upd As Drs
    Dim B As Drs
    Dim Sk$()
    Sk = Dbt_SkFny(Db, T)
    Set B = Dbt_Drs_Wh(Db, T)
Set Ins = Drs_XMinus(A, B, Sk)
Set Upd = Drs_XDif(A, B)
Drs_XIns_Dbt Ins, Db, T
Drs_XUpd_Dbt Upd, Db, T
End Sub
Function TT_Tny(TT) As String()
TT_Tny = FF_Fny(TT)
End Function
Function Fbtt_Wb(A, TT) As Workbook
Dim Wb As Workbook, T
    Set Wb = New_Wb
For Each T In TT_Tny(TT)
    Wb_XAdd_Fbt Wb, A, T
Next
Wb_XDlt_FstWs Wb
Set Fbtt_Wb = Wb
End Function

Sub Wb_XAdd_Fbt(A As Workbook, Fb, T)
Dim Ws As Worksheet
Dim A1 As Range
    Set Ws = Wb_XAdd_Ws(A, T, AtEnd:=True)
    Set A1 = Ws_A1(Ws)
Fbt_XPut_AtCell Fb, T, A1
End Sub
Sub Fbt_XPut_AtCell(A, T, AtCell As Range)

End Sub
Function Drs_XDif(A As Drs, B As Drs) As Drs

End Function
Sub Drs_XRpl_Dbt_Wh(A As Drs, Db As Database, T$, Optional BExpr$)
Const CSub$ = CMod & "DrsRplDbt"
Dim Miss$(), FnyDb$(), FnyDrs$()
FnyDb = Dbt_Fny(Db, T)
FnyDrs = A.Fny
Miss = AyMinus(FnyDb, FnyDrs)
If Sz(Miss) > 0 Then
    XThw CSub, "Some field in Drs is missing in Db-T", "Drs-Fny Dbt-Fny Missing Db] T", FnyDrs, FnyDb, Miss, Db_Nm(Db), T
'    XThw CSub, "Some field in Drs is missing in Db-T, Where [Drs-Fny] [Dbt-Fny] [Missing] [Db] [T]", FnyDrs, FnyDb, Db_Nm(Db), T
    'XThw CSub DrsRplDbt", "[Db]-[T]-[Fny] has [Missing-Fny] according to [Drs-Fny]", Db_Nm(Db), T, FnyDb, Miss, FnyDrs
End If
Db.Execute QDlt_Fm_OWh(T, BExpr)
Dim Dr, Rs As DAO.Recordset
Set Rs = Dbt_Rs(Db, T)
With Rs
    For Each Dr In Drs_XSel(A, Dbt_Fny(Db, T)).Dry
        Dr_XIns_Rs Dr, Rs
    Next
    .Close
End With
End Sub

Sub Drs_XUpd_Dbt(A As Drs, Db As Database, T)
Dim Sqy$(): GoSub X
DbSqy_XRun Db, Sqy
Exit Sub
X:
    Dim Dr, Fny$(), Dry(), Sk$()
    Fny = A.Fny
    Sk = Dbt_SkFny(Db, T)
    Dry = A.Dry
    For Each Dr In AyNz(Dry)
        Push Sqy, UpdSqlFmt(T, Sk, Fny, Dr)
    Next
    Return
End Sub


Private Sub Z()
Z_DtInsDbSqy
MDao__Dta:
End Sub
