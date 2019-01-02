Attribute VB_Name = "MIde_Mth_Fb_Tmp_Tbl"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Mth_Fb_Tmp_Tbl."
Private X_W As Database
Const IMthMch$ = "MthMch"
Const OMthMd$ = "MthMd"
Const YFny$ = "MthNm MdNm MchTyStr MchStr"

Sub XBrw_TmpTblMthMd()
Dbq_XBrw W, DbtFF_Sql(W, "#MthMd", "MdNm MthNm")
End Sub

Sub XBrw_TmpTblMthNy()
Dbt_XBrw W, "#MthNy"
End Sub

Private Property Get MthNy() As String()
MthNy = Dbt_FstStrCol(W, "#MthNy")
End Property

Sub XRfh_TmpTblMthMd()
XRfh_TmpMthNyTbl
Const T$ = "#MthMd"
Dbt_XDrp W, T
W.Execute "Create Table [#MthMd] (MthNm Text Not Null, MdNm Text(31),MthMchStr Text)"
'W.Execute CrtSkSql(T, "MthNm")
'Ay_XIns_Dbt MthNy, W, T ' MthNy is from #MthNy
Stop '
XUpd
End Sub

Sub XRfh_TmpMthNyTbl()
Const T$ = "#MthNy"
Dbt_XDrp W, T
W.Execute "Create Table [#MthNy] (MthNm Text)"
'W.Execute CrtSkSql(T, "MthNm")
'Ay_XIns_Dbt Md_MthNy(Md("AAAMod")), W, T
Stop '
End Sub

Private Property Get W() As Database
If IsNothing(X_W) Then Set X_W = MthDb
Set W = X_W
End Property

Private Sub WDrp_TmpTblMthNy()
Dbt_XDrp W, "#MthNy"
End Sub

Private Property Get WMchDic() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = Dbq_Dic(W, "Select MthMchStr,ToMdNm from MthMch order by Seq,MthMchStr")
Set WMchDic = X
End Property

Private Function WMd_MthNy(A As CodeModule) As String()
WMthNy_XEns_Cache A
WMd_MthNy = Dbtf_StrCol(W, "#MthNy", "MthNm")
End Function

Private Sub WMthNy_XEns_Cache(A As CodeModule)
Const T$ = "#MthNy"
If Dbt_Exist(W, T) Then Exit Sub
W.Execute "Create Table [#MthNy] (MthNm Text(255) Not Null)"
'W.Execute CrtSkSql(T, "MthNm")
'Ay_XIns_Dbt Md_MthNy(A), W, T
End Sub

Private Function XDr(MthNm, XDic As Dictionary) As Variant()
With StrDicMch(MthNm, XDic)
    If .Patn = "" Then
        XDr = Array(MthNm, "", "AAAMod")
    Else
        XDr = Array(MthNm, .Patn, .Rslt)
    End If
End With
End Function

Sub XUpd()
Const CSub$ = CMod & "XUpd"
Dim Rs As DAO.Recordset, XDic As Dictionary
Set XDic = Dbq_Dic(W, "Select MthMchStr,ToMdNm from MthMch order by Seq Desc,Ty")
Set Rs = Dbq_Rs(W, "Select MthNm,MthMchStr,MdNm from [#MthMd] where IIf(IsNull(MthMchStr),'',MthMchStr)=''")
While Not Rs.EOF
    Dr_XUpd_Rs XDr(Rs.Fields("MthNm").Value, XDic), Rs
    Rs.MoveNext
Wend
Dim A%, B%, C%
A = Dbt_NRec(W, "#MthNy")
B = Dbt_NRec(W, "#MthMd")
C = Dbt_NRec(W, "#MthMd", "MdNm='AAAMod'")
C = Dbq_Val(W, "Select count(*) from [#MthMd] where MdNm='AAAMod'")
Debug.Print CSub, "A: #MthNy-Cnt "; A
Debug.Print CSub, "B: #MthMd-Cnt "; B
Debug.Print CSub, "C: #MthMd-Wh-MdNm=AAAMod-Cnt "; C
Dbq_XBrw MthDb, "select * from [#MthMd] where MthMchStr='' order by MthNm"
End Sub

Private Property Get ZZMd() As CodeModule
Set ZZMd = Md("AAAMod")
End Property

Private Property Get ZZMthNy() As String()
ZZMthNy = WMd_MthNy(ZZMd)
End Property

Private Sub Z_ZZMthNy()
Brw ZZMthNy
End Sub

Private Sub ZZ()
XRfh_TmpTblMthMd
XUpd
End Sub

Private Sub Z()
Z_ZZMthNy
End Sub
