Attribute VB_Name = "MXls_Dao"
Option Compare Binary
Option Explicit

Sub Db_OupWb(A As Database, Wb As Workbook, OupNmSsl$)
'OupNm is used for Table-Name-@*, Ws_CdNm-Ws*, LoNm-Tbl*
Dim Ay$(), OupNm
Ay = Ssl_Sy(OupNmSsl)
Wb_XHas_OupNy_XAss Wb, Ay
Dim T$
For Each OupNm In Ay
    T = "@" & OupNm
    'Db_OupWb A, T, Wb, OupNm
Next
End Sub

Sub Db_Wb_ForAllOupTbl(A As Database, Wb As Workbook, OupNm)
'OupNm is used for Ws_CdNm-Ws*, LoNm-Tbl*
Dim Ws As Worksheet
Set Ws = Wb_Ws_Cd(Wb, "WsO" & OupNm)
'Dbt_XPut_Ws A, T, Ws
End Sub

Function Db_Wb_ForAllTmpTbl(A As Database, Optional UseWc As Boolean) As Workbook
Set Db_Wb_ForAllTmpTbl = Dbtt_Wb(A, Db_TmpTny(A), UseWc)
End Function

Function Dbt_XPut_AtCell1(A As Database, T, At As Range) As ListObject
Set Dbt_XPut_AtCell1 = Lo_XSet_Nm(Sq_Lo_RPL(Dbt_Sq(A, T), At), TblNm_LoNm(T))
End Function

Function Dbt_XPut_AtCell(A As Database, T, At As Range) As Range
Set Dbt_XPut_AtCell = Sq_XPut_AtCell(Dbt_Sq(A, T), At)
End Function

Function Dbt_XPut_AtCell_BY_CN(A As Database, T, At As Range, Optional LoNm0$) As ListObject
If XTak_FstChr(T) <> "@" Then Stop
Dim LoNm$, Lo As ListObject
If LoNm0 = "" Then
    LoNm = "Tbl" & XRmv_FstChr(T)
Else
    LoNm = LoNm0
End If
Dim AtA1 As Range, CnStr, Ws As Worksheet
Set AtA1 = RgRC(At, 1, 1)
Set Ws = Rg_Ws(At)
With Ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        QQ_Fmt("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share D", A.Name) _
        , _
        "eny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Databa" _
        , _
        "se Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Je" _
        , _
        "t OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Com" _
        , _
        "pact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=" _
        , _
        "False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        ), Destination:=AtA1).QueryTable '<---- At
        .CommandType = xlCmdTable
        .CommandText = Array(T) '<-----  T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnINF = True
        .ListObject.DisplayName = LoNm '<------------ LoNm
        .Refresh BackgroundQuery:=False
    End With

End Function

Function Dbt_XPut_Fx(A As Database, T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
Dim O As Workbook, Ws As Worksheet
Set O = Fx_Wb(Fx)
Set Ws = Wb_Ws(O, WsNm)
Ws_ClrLo Ws
Stop ' LoNm need handle?
Dbt_XPut_Ws A, T, Wb_Ws(O, WsNm)
Set Dbt_XPut_Fx = O
End Function

Sub Dbt_XPut_Ws(A As Database, T, Ws As Worksheet)
'Assume the Ws_CdNm is WsXXX and there will only 1 Lo with Name TblXXX
'Else stop
Dim Lo As ListObject
Set Lo = Ws_FstLo(Ws)

If Not XHas_Pfx(Ws.CodeName, "WsO") Then Stop
If Ws.ListObjects.Count <> 1 Then Stop
If Mid(Lo.Name, 4) <> Mid(Ws.CodeName, 4) Then Stop
Dbt_XRpl_Lo A, T, Lo
End Sub

Function Dbt_XRpl_Lo1(A As Database, T, At As Range, Optional UseWc As Boolean) As ListObject
Dim N$, Q As QueryTable
N = TblNm_LoNm(T)
Set Dbt_XRpl_Lo1 = Lo_XSet_Nm(Rg_Lo(Dbt_XPut_AtCell(A, T, At)), N)
End Function

Function Dbt_XRpl_Lo2(A As Database, T, Lo As ListObject, Optional ReSeqSpec$) As ListObject
Set Dbt_XRpl_Lo2 = Sq_Lo_RPL(Dbt_Sq_RESEQ(A, T, ReSeqSpec), Lo)
End Function

Function Dbt_XRpl_Lo(A As Database, T, Lo As ListObject) As ListObject
Dim Sq(), Drs As Drs, Rs As DAO.Recordset
Set Rs = Dbt_Rs(A, T)
If Not Ay_IsEq(Rs_Fny(Rs), Lo_Fny(Lo)) Then
    Debug.Print "--"
    Debug.Print "Rs"
    Debug.Print "--"
    Ay_XDmp Rs_Fny(Rs)
    Debug.Print "--"
    Debug.Print "Lo"
    Debug.Print "--"
    Ay_XDmp Lo_Fny(Lo)
    Stop
End If
Sq = Sq_Add_SngQuote(Rs_Sq(Rs))
Lo_XMin Lo
Sq_XPut_AtCell Sq, Lo.DataBodyRange
Set Dbt_XRpl_Lo = Lo
End Function

Sub Dbtt_Fx(A As Database, TT, Fx$)
Dbtt_Wb(A, TT).SaveAs Fx
End Sub

Function Dbtt_Wb(A As Database, TT, Optional UseWc As Boolean) As Workbook
Dim O As Workbook
Set O = New_Wb
Set Dbtt_Wb = Wb_XAdd_Dbtt(O, A, TT, UseWc)
Wb_Ws(O, "Sheet1").Delete
End Function

Sub Fbt_XPut_AtCell_ByWbCn(A As Database, T, AtCell As Range)
Dim Q As QueryTable
Set Q = Rg_Ws(AtCell).ListObjects.Add(SourceType:=0, Source:=Fb_CnStr_Ado(A), Destination:=AtCell).QueryTable
Qt_XSet_TblNm Q, T
End Sub

Sub Qt_XSet_StdPm(A As QueryTable, T)
With A
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = T
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub Qt_XSet_TblNm(A As QueryTable, TblNm)
With A
    .CommandType = xlCmdTable   '<==
    .CommandText = TblNm            '<==
End With
Qt_XSet_StdPm A, TblNm
End Sub

Sub TT_Fx(TT$, Fx$)
Dbtt_Fx CurDb, TT, Fx
End Sub

Function Tbl_Ws(T, Optional WsNm$ = "Data") As Worksheet
Set Tbl_Ws = Lo_Ws(Sq_Lo(Tbl_Sq(T), New_A1(WsNm)))
End Function

Private Sub ZZ()
Dim A As Database
Dim B As Workbook
Dim C$
Dim D
Dim E As ListObject
Dim F As Range
Dim G As Boolean
Dim H As QueryTable
Dim I As Worksheet
Dim XX
Db_OupWb A, B, C
Db_Wb_ForAllOupTbl A, B, D
Db_Wb_ForAllTmpTbl A, G
Dbt_XPut_AtCell A, D, F
Dbt_XPut_AtCell1 A, D, F
Dbt_XPut_AtCell_BY_CN A, D, F, C
Dbt_XPut_Fx A, D, C, C, C
Dbt_XPut_Ws A, D, I
Dbt_XRpl_Lo A, D, E
Dbt_XRpl_Lo1 A, D, F, G
Dbt_XRpl_Lo2 A, D, E, C
Dbtt_Fx A, D, C
Dbtt_Wb A, D, G
Fbt_XPut_AtCell_ByWbCn A, D, F
Qt_XSet_StdPm H, D
Qt_XSet_TblNm H, D
TT_Fx C, C
Tbl_Ws D, C
End Sub

'Function Tbl_XLnk_Fx(T, Fx$, Optional WsNm$ = "Sheet1") As String()
'Tbl_XLnk_Fx = Dbt_XLnk_Fx(CurDb, T, Fx, WsNm)
'End Function
'Function TblPutAt(T, At As Range) As Range
'Set TblPutAt = Dbt_XPut_AtCell(CurDb, T, At)
'End Function
'Function TblPutFx(T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
'Set TblPutFx = Dbt_XPut_Fx(CurDb, T, Fx, WsNm, LoNm)
'End Function
'Function TblRg(T, At As Range) As Range
'Set TblRg = Dbt_XPut_AtCell(CurDb, T, At)
'End Function
'Sub DbLnkFx(A As Database, Fx$, WsNy0)
'Dim Ws
'For Each Ws In CvNy(WsNy0)
'   Dbt_XLnk_Fx A, Fx, Ws
'Next
'End Sub
'Function DbInfWb(A As Database) As Workbook
'Set DbInfWb = Ds_Wb(DbInfDs(A))
'End Function
'Function TTWb(TT0, Optional UseWc As Boolean) As Workbook
'Set TTWb = Dbtt_Wb(CurDb, TT0, UseWc)
'End Function
'Sub TTWbBrw(TT, Optional UseWc As Boolean)
'Wb_XVis TTWb(TT, UseWc)
'End Sub
'Sub TtWrtFx(TT, Fx$)
'Dbtt_WrtFx CurDb, TT, Fx
'End Sub
Private Sub Z()
End Sub
