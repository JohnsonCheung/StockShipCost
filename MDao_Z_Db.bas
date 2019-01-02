Attribute VB_Name = "MDao_Z_Db"
Option Compare Binary
Option Explicit

Function Db_XAdd_Td(A As Database, Td As DAO.TableDef) As DAO.TableDef
A.TableDefs.Append Td
Set Db_XAdd_Td = Td
End Function

Sub Db_XAdd_TmpTbl(A As Database)
Db_XAdd_Td CurDb, TmpTd
End Sub

Sub Db_XApp_Td(A As Database, Td As DAO.TableDef)
A.TableDefs.Append Td
End Sub

Sub Db_XBrw(A As Database)
Fb_XBrw A.Name
End Sub

Function Db_XChk__Idx(A As Database) As String()
Dim T$()
T = Ay_XSrt(Db_Tny(A))
Db_XChk__Idx = Ay_XAlign_1T(AyAp_XAdd(Db_XChk_Pk(A, T), Db_XChk_Sk(A, T)))
End Function

Function Db_XChk_Pk(A As Database, Tny$()) As String()
Db_XChk_Pk = Ay_XExl_EmpEle(AyMapPX_Sy(Tny, "DbtMsgPk", A))
End Function

Function Db_XChk_Sk(A As Database, Tny$()) As String()
End Function

Function Db_CnStrAy(A As Database) As String()
Dim T$(), S()
T = Ay_XQuote_SqBkt(Db_Tny(A))
S = AyMapXP_Ay(T, "GetCnSTr_DB_T", A)
Db_CnStrAy = AyAB_Ly_NON_EMP_B(T, S)
End Function

Sub Db_XCrt_Qry(A As Database, Q, Sql)
If Not DbHasQry(A, Q) Then
    Dim QQ As New QueryDef
    QQ.Sql = Sql
    QQ.Name = Q
    A.QueryDefs.Append QQ
Else
    A.QueryDefs(Q).Sql = Sql
End If
End Sub

Sub DbCrtResTbl(A As Database)
Dbt_XDrp A, "Res"
DoCmd.RunSQL "Create Table Res (ResNm Text(50), Att Attachment)"
End Sub

Sub DbCrtTbl(A As Database, T$, FldDclAy)
A.Execute QQ_Fmt("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function DbDesy(A As Database) As String()
Dim T$(), D$()
T = Db_Tny(A)
DbDesy = Ay_XExl_EmpEle(AyMapPX_Sy(T, "DbtTbl_Des", A))
End Function

Sub Db_XDrp_AllTmpTbl(A As Database)
Dbtt_XDrp A, Db_TmpTny(A)
End Sub

Sub Db_XDrp_Qry(A As Database, Q)
If DbHasQry(A, Q) Then A.QueryDefs.Delete Q
End Sub

Function Db_Ds(A As Database, Optional Tny0, Optional DsNm$ = "Ds") As Ds
Dim DtAy() As Dt
    Dim U%, Tny$()
    Tny = DftTny(Tny0, A.Name)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        Set DtAy(J) = Dbt_Dt(A, Tny(J))
    Next
Set Db_Ds = Ds(DtAy, DftDb_Nm(DsNm, A))
End Function

Sub Db_XEns_Tmp1Tbl(A As Database)
If Dbt_Exist(A, "Tmp1") Then Exit Sub
Dbq_XRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub

Function DbHasQry(A As Database, Q) As Boolean
DbHasQry = Dbq_XHas_Rec(A, QQ_Fmt("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Function Tbl_Exist(T) As Boolean
Tbl_Exist = Dbt_Exist(CurDb, T)
End Function

Function Dbt_Exist(A As Database, T) As Boolean
Dbt_Exist = Itr_XHas_Nm(Db_XReOpn(A).TableDefs, T)

'Q = QQ_Fmt("Select Name from MSysObjects where Type in (1,6) and Name='?'", T)
'Dbt_Exist = Rs_IsNoRec(A.OpenRecordset(Q))
End Function

Sub Dbt_Exist_XAss(A As Database, T, Fun$)
If Dbt_Exist(A, T) Then XThw Fun, "Db already have Tbl", "Db Tbl", Db_Nm(A), T
End Sub

Function Db_IsOk(A As Database) As Boolean
On Error GoTo X
Db_IsOk = IsStr(A.Name)
Exit Function
X:
End Function

Sub Db_XKill(A As Database)
Dim F$
F = A.Name
A.Close
Kill F
End Sub

Function Db_Nm$(A As Database)
Db_Nm = Obj_Nm(A)
End Function

Function DbOupTny(A As Database) As String()
DbOupTny = Dbq_Sy(A, "Select Name from MSysObjects where Name like '@*' and Type =1")
End Function

Function DbPth$(A As Database)
DbPth = Ffn_Pth(A.Name)
End Function

Function DbQny(A As Database) As String()
DbQny = Dbq_Sy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Function DbQryRs(A As Database, Qry) As DAO.Recordset
Set DbQryRs = A.QueryDefs(Qry).OpenRecordset
End Function

Function DbReOpn(A As Database) As Database
Set DbReOpn = Fb_Db(A.Name)
End Function

Sub DbResClr(A As Database, ResNm$)
A.Execute "Delete From Res where ResNm='" & ResNm & "'"
End Sub

Function Db_TdSclLy(A As Database) As String()
Db_TdSclLy = Ay_Sy(AyOfAy_Ay(AyMap(ItrMap(A.TableDefs, "Td_SclLy"), "Td_SclLy_XAdd_Pfx")))
End Function

Sub DbSetTDes(A As Database, TDes As Dictionary)
Stop '
'Dbt_Des(A, B.T) = B.Des
End Sub

Sub DbSqy_XRun(A As Database, Sqy$())
Dim Q
For Each Q In AyNz(Sqy)
   A.Execute Q
Next
End Sub

Function Db_SrcTny(A As Database) As String()
Dim S$()
Dim T$(), I
T = Ay_XQuote_SqBkt(Db_Tny(A))
For Each I In AyNz(T)
    PushI S, Dbt_SrcTblNm(A, I)
Next
Db_SrcTny = AyAB_Ly_NON_EMP_B(T, S)
End Function
Function CurDb_TmpTny() As String()
CurDb_TmpTny = Db_TmpTny(CurDb)
End Function
Function TmpTny() As String()
TmpTny = CurDb_TmpTny
End Function
Function Db_TmpTny(A As Database) As String()
Db_TmpTny = Ay_XWh_Pfx(Db_Tny(A), "#")
End Function

Function Db_Tny(A As Database) As String()
Db_Tny = Db_Tny_DAO(A)
End Function

Function Db_Tny_ADO(A As Database) As String()
Db_Tny_ADO = Fb_Tny(A.Name)
End Function
Function Db_XReOpn(A As Database) As Database
Set Db_XReOpn = Fb_Db(A.Name)
End Function
Function Db_Tny_DAO(A As Database) As String()
Dim T As TableDef, O$()
Dim X As DAO.TableDefAttributeEnum
X = DAO.TableDefAttributeEnum.dbHiddenObject Or DAO.TableDefAttributeEnum.dbSystemObject
For Each T In Db_XReOpn(A).TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI Db_Tny_DAO, T.Name
    End Select
Next
End Function

Function Db_TnyMSYS(A As Database) As String()
Db_TnyMSYS = Dbq_Sy(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
End Function

Private Sub ZZ_DbQny()
Ay_XDmp DbQny(Fb_Db(Samp_Fb_Duty_Dta))
End Sub

Private Sub Z_Db_Ds()
Dim Db As Database, Tny0
Stop
ZZ1:
    Set Db = Fb_Db(Samp_Fb_Duty_Dta)
    Set Act = Db_Ds(Db)
    DsBrw CvDs(Act)
    Exit Sub
ZZ2:
    Tny0 = "Permit PermitD"
    Set Act = Db_Ds(CurDb, Tny0)
    Stop
End Sub

Private Sub Z_DbQny()
Ay_XDmp DbQny(CurDb)
End Sub

Private Sub ZZ()
Dim A As Database
Dim B As DAO.TableDef
Dim C$()
Dim D As Variant
Dim E$
Dim F As Drs
Dim G As Dictionary
Db_XAdd_Td A, B
Db_XAdd_TmpTbl A
Db_XApp_Td A, B
Db_XBrw A
Db_XChk_Pk A, C
Db_XChk_Sk A, C
Db_CnStrAy A
Db_XCrt_Qry A, D, D
DbCrtResTbl A
DbCrtTbl A, E, D
DbDesy A
Db_XDrp_Qry A, D
Db_Ds A, D, E
Db_XEns_Tmp1Tbl A
DbHasQry A, D
Dbt_Exist A, D
Dbt_Exist_XAss A, D, E
Db_IsOk A
Db_XKill A
Db_Nm A
DbOupTny A
DbPth A
DbQny A
DbQryRs A, D
DbReOpn A
DbResClr A, E
Db_TdSclLy A
DbSetTDes A, G
DbSqy_XRun A, C
Db_SrcTny A
Db_TmpTny A
Db_Tny A
Db_TnyMSYS A
End Sub

Private Sub Z()
Z_Db_Ds
Z_DbQny
End Sub
