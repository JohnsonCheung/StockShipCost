Attribute VB_Name = "MDao_Schm"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Schm."

Sub Db_XCrt_Schm(A As Database, Schm$)
Const CSub$ = CMod & "Db_XCrt_Schm"
Dim Ly$()
    Ly = Split(Schm, vbCrLf)

Dim E$()
    E = SchmLy_Er(Ly)
    If Sz(E) > 0 Then
        XThw CSub, "There is error in Schm", "Db Schm Er", Db_Nm(A), LinesAddIxPfx(Schm, BegFm:=1), E
    End If

Dim B As StruBase, I
B = New_StruBase(Ly)
For Each I In AyNz(Ay_XWh_T1SelRst(Ly, "Tbl"))
    DbTdStrB_XCrt A, I, B
Next
End Sub

Sub DbTdStrE_XCrt(A As Database, TdStr, EleVbl$)
DbTdStrB_XCrt A, TdStr, New_StruBaseEF(NewEF(NewDicVBL(EleVbl)))
End Sub

Sub TdStrE_XCrt(TdStr, EleVbl$)
DbTdStrE_XCrt CurDb, TdStr, EleVbl
End Sub

Sub DbTdStr_XCrt(A As Database, TdStr)
DbTdStrB_XCrt A, TdStr, EmpStruBase
End Sub

Sub TdStr_XCrt(A)
DbTdStr_XCrt CurDb, A
End Sub

Sub TdStrB_XCrt(A, B As StruBase)
DbTdStrB_XCrt CurDb, A, B
End Sub

Sub DbTdStrB_XCrt(A As Database, TdStr, B As StruBase)
Const CSub$ = CMod & "DbTdStrCrtB"
Dim Td As DAO.TableDef
Dim FDes As Dictionary
    Dim T$: T = RmvSfx(Lin_T1(TdStr), "*")
    Dim Fny$(): Fny = TdStrFny(TdStr)
    Dim Sk1$(): Sk1 = TdStrSk(TdStr)
    Set Td = NewTdEF(T, Fny, Sk1, B.EF)
    Set FDes = DesDicFldDesDic(B.Des, T, Fny)
Dbt_Exist_XAss A, T, CSub

A.TableDefs.Append Td                                   '<--
Dbt_Des(A, Td.Name) = CStr(DicVal(B.Des, T & "."))       '<--

Dim F
For Each F In FDes.Keys
    Dbtf_Des(A, Td.Name, F) = FDes(F)                    '<--
Next
End Sub

Sub DbTdStr_XEns(A As Database, TdStr)
DbTdStr_XEns_B A, TdStr, EmpStruBase
End Sub

Sub DbTdStr_XEns_B(A As Database, TdStr, B As StruBase)
ChkAss TdStrEF_XChk(TdStr, B.EF)
Dim S$
S = Dbt_Stru(A, Lin_T1(Stru))
If S = "" Then
    DbTdStrB_XCrt A, Stru, B
    Exit Sub
End If
If S = Stru Then Exit Sub
Dbt_XReStru A, Stru, B.EF
End Sub

Sub Dbtf_Add(A As Database, T, F, EF As EF)
If Dbt_Existf(A, T, F) Then Exit Sub
A.TableDefs(T).Fields.Append New_Fd_EF(F, T, EF)
End Sub

Sub DbtFFAdd(A As Database, T, FF, EF As EF)
Dim F
For Each F In CvNy(FF)
    Dbtf_Add A, T, F, EF
Next
End Sub

Private Sub TdStrEF_XAss(A, EF As EF)
ChkAss TdStrEF_XChk(A, EF)
End Sub

Function NewTdEF(T, Fny$(), Sk$(), EF As EF) As DAO.TableDef
'If Ay_XHas(Fny, T & "Id") Then OPk = CrtPkSql(T) Else OPk = ""
Dim FdAy() As DAO.Field2
Dim F
For Each F In Fny
    PushObj FdAy, New_Fd_EF(F, T, EF)  '<===
Next
Stop
'Set NewTdEF = NewTd(T, FdAy, Sk)
End Function

Private Sub Z_Db_XCrt_Schm()
Dim Schm$, Db As Database
Set Db = TmpDb
Schm = Z_Db_XCrt_Schm1
GoSub Tst
Exit Sub
Tst:
    Db_XCrt_Schm Db, Schm
    Ass Dbt_Exist(Db, "A")
    Ass Dbt_Exist(Db, "B")
    Ay_IsEq_XAss Dbt_Fny(Db, "A"), Ssl_Sy("AId ANm ADte AATy Loc Expr Rmk")
    Ay_IsEq_XAss Dbt_Fny(Db, "B"), Ssl_Sy("BId ANm BNm BDte")
    'Db_XBrw Db
    Stop
    Return
End Sub

Private Property Get Z_Db_XCrt_Schm1$()
Const A_1$ = "Tbl A *Id | *Nm     | *Dte AATy Loc Expr Rmk" & _
vbCrLf & "Tbl B *Id | AId *Nm | *Dte" & _
vbCrLf & "Ele Txt AATy" & _
vbCrLf & "Ele Loc Loc" & _
vbCrLf & "Ele Expr Expr" & _
vbCrLf & "Ele Mem Rmk" & _
vbCrLf & "Fld Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']" & _
vbCrLf & "Fld Expr Txt [Expr=Loc & 'abc']" & _
vbCrLf & "Des Tbl     A     AA BB " & _
vbCrLf & "Des Tbl     A     CC DD " & _
vbCrLf & "Des Fld     ANm   AA BB " & _
vbCrLf & "Des Tbl.Fld A.ANm TF_Des-AA-BB"

Z_Db_XCrt_Schm1 = A_1
End Property

Private Sub Z_DbTdStr_XCrt()
Dim Td As DAO.TableDef
Dim B As DAO.Field2
GoSub X_Td
GoSub Tst
Exit Sub
Tst:
    Tbl_XDrp Td.Name
    Debug.Print ObjPtr(Td)
    CurDb.TableDefs.Append Td
    Debug.Print ObjPtr(Td)
    Set B = CurDb.TableDefs("#Tmp").Fields("B")
    CurDb.TableDefs("#Tmp").Fields("B").Properties.Append CurDb.CreateProperty(C_Des, dbText, "ABC")
    Return
X_Td:
    Dim FdAy() As DAO.Field
    Set B = New_Fd("B", dbInteger)
    PushObj FdAy, New_Fd_ID("#Tmp")
    PushObj FdAy, New_Fd("A", dbInteger)
    PushObj FdAy, B
    Set Td = NewTd("#Tmp", FdAy, "A B")
    Return
End Sub

Private Sub ZZ()
Dim A As Database
Dim B$
Dim C As Variant
Dim D As StruBase
Dim E As EF
Dim F$()

Db_XCrt_Schm A, B
End Sub

Private Sub Z()
Z_Db_XCrt_Schm
End Sub
