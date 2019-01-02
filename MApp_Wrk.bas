Attribute VB_Name = "MApp_Wrk"
Option Compare Binary
Option Explicit
Private X_W As Database

Property Get W() As Database
If IsNothing(X_W) Then WEns: WOpn
Set W = X_W
End Property

Sub WBrw()
Db_XBrw W
End Sub

Sub WClr()
Dim T, Tny$()
Tny = WTny: If Sz(Tny) = 0 Then Exit Sub
For Each T In Tny
    WDrp T
Next
End Sub

Sub WCls()
On Error Resume Next
X_W.Close
Set X_W = Nothing
End Sub

Sub WCrt()
Fb_XCrt WFb
End Sub

Property Get WDaoCn() As DAO.Connection
'Set WDaoCn = FbDaoCn(WFb)
End Property

Property Get WDb() As Database
Set WDb = X_W
End Property

Sub WDrp(TT)
Dbtt_XDrp W, TT
End Sub

Sub WEns()
If Not WExist Then WCrt
End Sub

Property Get WExist() As Boolean
WExist = Ffn_Exist(WFb)
End Property

Property Get WFb$()
WFb = Apn_WFb(Apn)
End Property

Sub WGen(Fx$, ParamArray WbFmtrFunAp())
Msg_XSet_MainMsg "Export to Excel ....."
Dim Av()
Av = WbFmtrFunAp
Ffn_XCrt_ByTp Fx
Wb_XFmt Fx_XRfh(Fx, WFb), Av
End Sub

Sub WImp(T$, LnkColVbl$, Optional WhBExpr$)
If XTak_FstChr(T) <> ">" Then Stop
Dbt_XImp_ByLnkVblCol W, T, LnkColVbl, WhBExpr
End Sub

Sub WImpTbl(TT)
Dbtt_Imp W, TT
End Sub

Sub WIni()
TpImp
WCls
Ffn_XDltIfExist WFb
WOpn
End Sub

Sub WKill()
WCls
Ffn_XDltIfExist WFb
End Sub

Sub WOpn()
Fb_XEns WFb
Set X_W = Fb_Db(WFb)
End Sub

Property Get WOupWb() As Workbook
Set WOupWb = Fb_Wb_OUP_TBL(WFb)
End Property

Property Get WPth$()
WPth = ApnWPth(Apn)
End Property

Sub WPthBrw()
Pth_XBrw WPth
End Sub

Sub WReOpn()
WCls
WOpn
End Sub

Sub WRun(A)
On Error GoTo X
W.Execute A
Exit Sub
X:
Debug.Print Err.Description
Debug.Print A
Debug.Print "?WStru("""")"
On Error Resume Next
'Db_XCrt_Qry W, "Query1", A
Stop
End Sub


Sub WSetDb(A As Database)
Set X_W = A
End Sub

Property Get WTny() As String()
WTny = Ay_XSrt(Db_Tny(W))
End Property

Property Get WWbCnStr$()
WWbCnStr = Fb_CnStr_OleAdo(WFb)
End Property

Function Wt_XChk_Col_ByLnkColVbl(T$, LnkColVbl$) As String()
Wt_XChk_Col_ByLnkColVbl = Dbt_XChk_Col_ByLnkColVbl(W, T, LnkColVbl)
End Function

Sub Wt_XCrt_FldLisTbl(T, TarTbl$, KyFld$, Fld$, Sep$, Optional IsMem As Boolean, Optional LisFldNm0$)
Dbt_XCrt_FldLisTbl W, T, TarTbl, KyFld, Fld, Sep, IsMem, LisFldNm0
End Sub

Function Wt_Fny(T$) As String()
Wt_Fny = Dbt_Fny(W, T)
End Function

Function Wt_XLnk_Fb(T$, Fb$) As String()
Wt_XLnk_Fb = Dbt_XLnk_Fb(W, T, Fb)
End Function

Function Wt_XLnk_Fx(T$, Fx$, Optional WsNm$ = "Sheet1") As String()
Wt_XLnk_Fx = Dbt_XLnk_Fx(W, T, Fx, WsNm)
End Function

Function Wt_PrpLoFmtrVbl$(T$)
Wt_PrpLoFmtrVbl = Dbt_PrpLoFmtrVbl(WDb, T)
End Function

Function Wt_Stru$(T$)
Wt_Stru = Dbt_Stru(W, T)
End Function

Sub Wt_XImp(T$, ColLnk$())
Dbt_XImp W, T, ColLnk
End Sub

Sub Wt_XReSeq(T$, ReSeqSpec$)
Dbt_XReSeq_BY_SPEC W, T, ReSeqSpec
End Sub

Sub Wt_XRen(T$, ToTbl$)
Dbt_XRen W, T, ToTbl
End Sub

Sub Wt_XRen_Col(T$, Fm$, NewColNm$)
Dbt_XRen_Col W, T, Fm, NewColNm
End Sub

Sub Wtf_XAdd_Expr(T$, F$, Expr$)
Dbtf_XAdd_Expr W, T, F, Expr
End Sub

Sub Wtt_XLnk_Fb(TT$, Fb$, Optional Fbtt$)
Dbtt_XLnk_Fb W, TT, Fb$, Fbtt
End Sub

Function Wtt_Stru$(TT)
Wtt_Stru = Dbtt_Stru(W, TT)
End Function

Sub Wtt_Stru_XDmp(TT)
D Wtt_Stru(TT)
End Sub

Private Sub ZZ()
Dim A
Dim B$
Dim C()
Dim D As Database
Dim E As Boolean
Dim G$()
Dim XX
WBrw
WClr
WCls
WCrt
WDrp A
WEns
WGen B, C
WImp B, B, B
WImpTbl A
WIni
WKill
WOpn
WPthBrw
WReOpn
WRun A
WSetDb D
Wt_XChk_Col_ByLnkColVbl B, B
Wt_XCrt_FldLisTbl A, B, B, B, B, E, B
Wt_Fny B
Wt_XLnk_Fb B, B
Wtt_Stru A
Wtt_Stru_XDmp A
XX = W()
XX = WDb()
XX = WExist()
XX = WFb()
XX = WOupWb()
XX = WPth()
XX = WTny()
XX = WWbCnStr()
End Sub

Private Sub Z()
End Sub
