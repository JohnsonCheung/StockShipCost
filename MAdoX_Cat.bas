Attribute VB_Name = "MAdoX_Cat"
Option Compare Binary
Option Explicit
Private Function Cat_XHas_Tbl(A As Catalog, T) As Boolean
Cat_XHas_Tbl = Itr_XHas_Nm(A.Tables, T)
End Function

Private Function Cat_Tny(A As Catalog) As String()

Cat_Tny = Itr_Ny(A.Tables) 'Catalog-A for Fx will return XXX$ for WsNm & XXX for ListObjectName
                           'Eg, 2-Ws-Sheet1-&-Sheet-2 and 1-Listobject-Name-XXX ==> Sheet1$, Sheet2$, XXX
End Function

Private Function Fb_Cat(A) As Catalog
Set Fb_Cat = Cn_Cat(Fb_Cn(A))
End Function

Private Function Cn_Cat(A As ADODB.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set Cn_Cat = O
End Function

Private Function Fx_Cat(A) As Catalog
Set Fx_Cat = Cn_Cat(Fx_Cn(A))
End Function

Private Sub Z_Fb_Tny()
Ay_XDmp Fb_Tny(Samp_Fb_Duty_Dta)
End Sub

Private Sub Z_Fx_WsNy()
Ay_XDmp Fx_WsNy(Samp_Fx_KE24)
End Sub

Function FbTbl_Exist(A, T) As Boolean
FbTbl_Exist = Ay_XHas(Fb_Tny(A), T)
End Function
Function Fxw_Exist(A, W) As Boolean
Fxw_Exist = Ay_XHas(Fx_WsNy(A), W)
End Function

Sub Samp_Fb_Tny()
D Fb_Tny(Samp_Fb_Duty_Dta)
End Sub
Function Fb_Tny(A) As String()
Fb_Tny = Fb_Tny_ADO(A)
End Function

Function Fb_Tny_ADO(A) As String()
Fb_Tny_ADO = Ay_XExl_Likss(Cat_Tny(Fb_Cat(A)), "MSys* f_*_Data")
End Function
Function FxCatTny_WsNy(A$()) As String()
Dim I, B$
For Each I In AyNz(A)
    B = XRmv_SngQuote(I)
    If XHas_Sfx(B, "$") Then
        PushI FxCatTny_WsNy, XRmv_LasChr(B)
    End If
Next
End Function
Function Fx_CatTny(A) As String()
Fx_CatTny = Cat_Tny(Fx_Cat(A))
End Function
Function Fx_WsNy(A) As String()
Fx_WsNy = FxCatTny_WsNy(Fx_CatTny(A))
End Function
Private Sub Z_Fxw_Fny()
Dim W
For Each W In Fx_WsNy(Samp_Fx_KE24)
    D W & "<====================="
    D Fxw_Fny(Samp_Fx_KE24, W)
Next
End Sub
Private Function FbtCatTbl(A, T) As ADOX.Table
Set FbtCatTbl = Fb_Cat(A).Tables(T)
End Function
Function Fbt_Fny(A, T) As String()
Fbt_Fny = Itr_Ny(Fb_Cat(A).Tables(T).Columns)
End Function
Function WsNm_CatTblNm$(A)
Dim O$
    O = A & "$"
If XHas_Spc(A) Then O = XQuote_Sng(O)
WsNm_CatTblNm = O
End Function

Private Function Fxw_CatTbl(A, W) As ADOX.Table
Set Fxw_CatTbl = Fx_Cat(A).Tables(WsNm_CatTblNm(W))
End Function
Private Function Fxt_CatTbl(A, CatTblNm) As ADOX.Table
Set Fxt_CatTbl = Fx_Cat(A).Tables(CatTblNm)
End Function
Function Fxw_Fny(A, W) As String()
'Using WsNm-W will not return any Fny
'Only CatTblNm will return Fny, use Fxt_Fny

Dim T$
    T = WsNm_CatTblNm(W)
Fxw_Fny = Itr_Ny(Fx_Cat(A).Tables(T).Columns)
End Function

Function Fxt_Fny(A, CatTblNm) As String()
Fxt_Fny = Itr_Ny(Fx_Cat(A).Tables(CatTblNm).Columns)
End Function

Function Fx_Fny_FstWs(A) As String()
Fx_Fny_FstWs = Fxw_Fny(A, Fx_FstWsNm(A))
End Function

Private Sub Z()
Z_Fb_Tny
Z_Fxw_Fny
Z_Fx_WsNy
MAdoX_Cat:
End Sub
