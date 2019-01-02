Attribute VB_Name = "MDao_Chk_Col"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Chk_Col."

Function Dbt_XChk_Col_ByLnkColVbl(A As Database, T, LnkColVbl$) As String()
Dim Fny$()
Dim SemiColonDaoShtTyStrAy$()
    Dim Ay() As LnkCol
    Ay = LnkColVbl_LnkColAy(LnkColVbl)
    Fny = LnkColAy_ExtNy(Ay)
    SemiColonDaoShtTyStrAy = LnkColAy_SemiColonDaoShtTyStrAy(Ay)
Dim O$()
O = Dbt_XChk_Fny(A, T, Fny) '<==============================
If Sz(O) > 0 Then
    Dbt_XChk_Col_ByLnkColVbl = O
    Exit Function
End If
Dbt_XChk_Col_ByLnkColVbl = Dbt_XChk_Ty(A, T, Fny, SemiColonDaoShtTyStrAy) '<=================
End Function

Function Dbt_XChk_Fny(A As Database, T, EptFny$()) As String()
Const CSub$ = CMod & "Dbt_XChk_Fny"
Dim Miss$(), ActFny$()
ActFny = Dbt_Fny(A, T)
Miss = AyMinus(EptFny, ActFny)
If Sz(Miss) = 0 Then Exit Function

Const M$ = "Some fields are missing in given table"
Const Ny0$ = "Db [Given Tbl] [Given table linking to file] [Linking to Tbl/Ws] [Missing fields] [Expected fields] [Given table fields]"
Dbt_XChk_Fny = FunMsgNyAp_Ly(CSub, M, Ny0, _
    Db_Nm(A), T, Dbt_LnkingFfn(A, T), Dbt_LinkingTblOrWs(A, T), Ay_XQuote_SqBkt_IfNeed(Miss), Ay_XQuote_SqBkt_IfNeed(EptFny), Ay_XQuote_SqBkt_IfNeed(ActFny))
End Function

Function Dbt_XChk_Ty(A As Database, T, ExtNy$(), SemiColonDaoShtTyStrAy$()) As String()
Const CSub$ = CMod & "Dbt_XChk_Ty"
Dim EptDaoTyAy() As DAO.DataTypeEnum
Dim SemiColonDaoShtTyStr$, J%
Dim F, M$, O$()
Dim ActTy As DAO.DataTypeEnum
For Each F In ExtNy
    SemiColonDaoShtTyStr = SemiColonDaoShtTyStrAy(J)
    EptDaoTyAy = DaoSemiColonTyStr_DaoTyAy(SemiColonDaoShtTyStr)
    ActTy = A.TableDefs(T).Fields(F).Type
    If Not Ay_XHas(EptDaoTyAy, ActTy) Then
        M = QQ_Fmt("Table[?] field[?] should have type[?], but now it has type[?]", T, F, SemiColonDaoShtTyStr, DaoTy_DaoShtTyStr(ActTy))
        PushI O, M
    End If
    J = J + 1
Next
If Sz(O) = 0 Then Exit Function
Dim NN$
M = "Some fields has unexpected data type"
NN = "Unexpect-Data-Type-Fields Db [Given Tbl] [Given table linking to file] [Linking to Tbl/Ws] [Given table fields]"
Dbt_XChk_Ty = FunMsgNyAp_Ly(CSub, M, NN, O, Db_Nm(A), T, Dbt_LnkingFfn(A, T), Dbt_LinkingTblOrWs(A, T), Ay_XQuote_SqBkt_IfNeed(Dbt_Fny(A, T)))
End Function

Private Sub ZZ()
Dim A As Database
Dim B
Dim C$
Dim D$()
Dim XX
Dbt_XChk_Col_ByLnkColVbl A, B, C
Dbt_XChk_Fny A, B, D
Dbt_XChk_Ty A, B, D, D
End Sub

Private Sub Z()
End Sub
