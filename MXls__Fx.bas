Attribute VB_Name = "MXls__Fx"
Option Compare Binary
Option Explicit
Const CMod$ = "MXls__Fx."

Function Fx_XEns$(A$)
If Not Ffn_Exist(A) Then Fx_XCrt A
Fx_XEns = A
End Function

Sub Fx_XCrt(A)
Wb_XSavAs(New_Wb, A).Close
End Sub

Function Fx_Fny(A$, Optional WsNm$ = "Sheet1") As String()
Fx_Fny = Itr_Ny(Fx_Cat(A).Tables(WsNm & "$").Columns)
End Function

Function Fx_FstWsNm$(A)
Fx_FstWsNm = Itr_FstItm(Fx_WsNy(A))
End Function

Function Fx_CnStr_Ole$(A)
Fx_CnStr_Ole = "OLEDb;" & Fx_CnStr_Ado(A)
End Function

Sub Fx_XBrw(A)
If Not Ffn_Exist(A) Then
    MsgBox "Excel file not found: " & vbCrLf & vbCrLf & A
    Exit Sub
End If
Dim C$
C = QQ_Fmt("Excel ""?""", A)
Shell C, vbMaximizedFocus
End Sub

Sub Fx_XRmv_WsIfExist(A, WsNm)
If Fxw_Exist(A, WsNm) Then
   Dim B As Workbook: Set B = Fx_Wb(A)
   Wb_Ws(B, WsNm).Delete
   Wb_XSav B
   Wb_XCls_NoSav B
End If
End Sub

Function Fxw_ARs(A, W, Optional F = 0) As ADODB.Recordset
Dim T$
T = W & "$"
If F = 0 Then
    Q = QSel_Fm(T)
Else
    Q = QSel_FF_Fm(F, T)
End If
Set Fxw_ARs = Cnq_ARs(Fx_Cn(A), Q)
End Function

Function Fx_TmpDb(Fx$, Optional WsNy0) As Database
Dim O As Database
   Set O = TmpDb
   Db_XLnk_Fx O, Fx, WsNy0
Set Fx_TmpDb = O
End Function

Function Fx_Wb(A) As Workbook
Const CSub$ = CMod & "Fx_Wb"
Set Fx_Wb = Xls_Wb_ByFx(Xls, A)
If IsSomething(Fx_Wb) Then Exit Function
Ffn_XAss_Exist A, CSub
Set Fx_Wb = Xls.Workbooks.Open(A)
End Function

Function Fx_Ws(A, Optional WsNm$ = "Data") As Worksheet
Set Fx_Ws = Wb_Ws(Fx_Wb(A), WsNm)
End Function

Function Fxw_Dt(A, Optional WsNm0$) As Dt
Dim N$: N = FxDftWsNm(A, WsNm0)
Dim Sql$: Sql = QQ_Fmt("Select * from [?$]", N)
Set Fxw_Dt = Drs_Dt(Fxq_Drs(A, Sql), N)
End Function

Function Fxw_IntAy_FstCol(A, W, Optional F = 0) As Integer()
Fxw_IntAy_FstCol = ARs_IntAy(Fxw_ARs(A, W, F))
End Function

Function Fxw_Sy_FstCol(A, W, Optional F = 0) As String()
Fxw_Sy_FstCol = ARs_Sy(Fxw_ARs(A, W, F))
End Function

Function Fx_WsCdNy(A) As String()
Dim Wb As Workbook
Set Wb = Fx_Wb(A)
Fx_WsCdNy = Wb_Ws_CdNy(Wb)
Wb.Close False
End Function

Function Fxq_Drs(A, Sql) As Drs
Set Fxq_Drs = ARs_Drs(Fx_Cn(A).Execute(Sql))
End Function

Sub Fxq_XRun(A, Sql)
Fx_Cn(A).Execute Sql
End Sub

Function Fxw_TyAy(A, W) As String()
Fxw_TyAy = CatTbl_TyAy(CatTbl(Fx_Cat(A), W))
End Function


Private Sub Z_Fx_FstWsNm()
Dim Fx$
Fx = Samp_Fx_KE24
Ept = "Sheet1"
GoSub T1
Exit Sub
T1:
    Act = Fx_FstWsNm(Fx)
    C
    Return
End Sub

Private Sub Z_Fx_TmpDb()
Dim Db As Database: Set Db = Fx_TmpDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Ay_XDmp Db_Tny(Db)
Db.Close
End Sub

Private Sub Z_Fx_WsNy()
Dim Fx$
'GoTo ZZ
GoSub T1
Exit Sub
T1:
    Fx = Samp_Fx_KE24
    Ept = Ssl_Sy("Sheet1 Sheet21")
    GoSub Tst
    Return
Tst:
    Act = Fx_WsNy(Fx)
    C
    Return
ZZ:
    Ay_XDmp Fx_WsNy(Samp_Fx_KE24)
    Const Fx1$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
    D Fx_WsNy(Fx1)
End Sub

Private Sub Z()
Z_Fx_FstWsNm
Z_Fx_TmpDb
Z_Fx_WsNy
End Sub
