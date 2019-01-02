Attribute VB_Name = "MDao_Lnk_Tbl_Fx"
Option Compare Binary
Option Explicit
Sub Db_XLnk_Fx(A As Database, Fx$, Optional WsNy0)
Dim W
For Each W In DftWsNy(WsNy0, Fx)
   Dbt_XLnk_Fx A, W, Fx, W
Next
End Sub

Function Fxw_Chk(A$, Optional WsNy0 = "Sheet1", Optional FxKind$ = "Excel file") As String()
If Ffn_NotExist(A) Then Fxw_Chk = Ffn_Msg_NotExist(A, FxKind): Exit Function
If Fxw_Exist(A, WsNy0) Then Exit Function
Dim M$
M = QQ_Fmt("? does not have expected worksheet", FxKind)
Fxw_Chk = FunMsgNyAp_Ly(CSub, M, "Folder File-Name Expected-Worksheet Worksheets-in-file", Ffn_Pth(A), Ffn_Fn(A), CvNy(WsNy0), Fx_WsNy(A))
End Function

Function Dbt_XLnk_Fx(A As Database, T, Fx$, Optional WsNm = "Sheet1") As String()
Dim O$(): O = Fxw_Chk(Fx, WsNm)
If Sz(O) > 0 Then Dbt_XLnk_Fx = O: Exit Function
Dim Cn$: Cn = Fx_CnStr_DAO(Fx)
Dim Src$: Src = WsNm & "$"
Dbt_XLnk_Fx = Dbt_XLnk(A, T, Src, Cn)
End Function
