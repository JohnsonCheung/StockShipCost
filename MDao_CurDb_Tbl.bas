Attribute VB_Name = "MDao_CurDb_Tbl"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_CurDb_Tbl."

Sub Tbl_XRen_XAdd_Pfx(T, Pfx$)
Dbt_XAdd_Pfx CurDb, T, Pfx
End Sub

Sub TblAp_XDrp(ParamArray TblAp())
Dim Av(): Av = TblAp
Dim T
For Each T In Av
    Tbl_XDrp T
Next
End Sub

Property Get TblAttDes$()
TblAttDes = Dbt_Des(CurDb, "Att")
End Property

Property Let TblAttDes(Des$)
Dbt_Des(CurDb, "Att") = Des
End Property

Property Get Tbl_Des$(T)
Tbl_Des = Dbt_Des(CurDb, T)
End Property

Property Let Tbl_Des(T, Des$)
Dbt_Des(CurDb, T) = Des
End Property

Sub Tbl_XDrp(T)
Dbt_XDrp CurDb, T
End Sub

Function Tbl_XChk_Col_ByColLnkVbl(T, ColLnkVbl$) As String()
Tbl_XChk_Col_ByColLnkVbl = Dbt_XChk_Col_ByLnkColVbl(CurDb, T, ColLnkVbl)
End Function

Function Tbl_Fny(A) As String()
Tbl_Fny = Dbt_Fny(CurDb, A)
End Function

Function Td_XHas_FldNm(T As TableDef, F) As Boolean
Td_XHas_FldNm = Fds_XHas_FldNm(T.Fields, F)
End Function

Function New_TblImpSpec(T, LnkSpec$, Optional WhBExpr$) As TblImpSpec
Dim O As New TblImpSpec
Set New_TblImpSpec = O.Init(T, LnkSpec$, WhBExpr)
End Function

Function Tbl_IsSys(T) As Boolean
Tbl_IsSys = Dbt_IsSys(CurDb, T)
End Function

Function Tbl_XLnk_Fb(T, Fb$, Optional Fbt$) As String()
Tbl_XLnk_Fb = Dbt_XLnk_Fb(CurDb, T, Fb$, Fbt)
End Function

Function Tbl_NCol&(T)
Tbl_NCol = Dbt_NCol(CurDb, T)
End Function

Function Tbl_NRec&(A)
Tbl_NRec = Sql_Lng(QQ_Fmt("Select Count(*) from [?]", A))
End Function

Function Tbl_NRow&(T, Optional WhBExpr$)
Tbl_NRow = Dbt_NRec(CurDb, T, WhBExpr)
End Function
Sub Tbl_XBrw(T)
Rs_XBrw Tbl_Rs(T)
End Sub

Sub Tbl_XOpn(T)
DoCmd.OpenTable T
End Sub

Function Tbl_Pk(T) As String()
Tbl_Pk = Dbt_Pk(CurDb, T)
End Function

Sub XCpy_TblPm_Fm_C_or_N(IsDev As Boolean)
CurDb.Execute "Delete * from Pm"
If IsDev Then
    CurDb.Execute "insert into Pm select * from Pm_C"
Else
    CurDb.Execute "insert into Pm select * from Pm_N"
End If
End Sub

Function Tbl_Rs(T) As DAO.Recordset
Set Tbl_Rs = Dbt_Rs(CurDb, T)
End Function

Function Tbl_SclLy(T) As String()
Tbl_SclLy = Dbt_SclLy(CurDb, T)
End Function

Function Tbl_SkFny(T) As String()
'Sk is secondary key.  Same name as the table and is unique
'Thw if there is key with same name as T, but not primary key.
'This is done as Dbt_SIdx
Tbl_SkFny = Dbt_SkFny(CurDb, T)
End Function

Function Tbl_Sq(T) As Variant()
Tbl_Sq = Dbt_Sq(CurDb, T)
End Function

Function Tbl_Src$(T)
Tbl_Src = CurDb.TableDefs(T).SourceTableName
End Function

Function Tbl_Stru$(T)
Tbl_Stru = Dbt_Stru(CurDb, T)
End Function

Sub Tbl_Stru_XDmp(T)
D Tbl_Stru(T)
End Sub

Property Get TF_Des$(T$, F$)
TF_Des = Dbtf_Des(CurDb, T, F)
End Property

Property Let TF_Des(T$, F$, V$)
Dbtf_Des(CurDb, T, F) = V
End Property

Private Sub ZZ_Tbl_Fny()
Ay_XDmp Tbl_Fny(">KE24")
End Sub

Private Sub ZZ_TblSq()
Dim A()
A = Tbl_Sq("@Oup")
Stop
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C()
Dim D$()
Dim E() As DAO.DataTypeEnum
Dim F As Boolean
Dim G As TableDef
Tbl_XRen_XAdd_Pfx A, B
TblAp_XDrp C
Tbl_Fny A
Td_XHas_FldNm G, A
New_TblImpSpec A, B, B
Tbl_Exist A
Tbl_IsSys A
Tbl_XLnk_Fb A, B, B
Tbl_NCol A
TblNm_LoNm A
Tbl_NRec A
Tbl_NRow A, B
End Sub

Private Sub Z()
End Sub
