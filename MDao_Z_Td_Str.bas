Attribute VB_Name = "MDao_Z_Td_Str"
Option Compare Binary
Option Explicit

Private Function Fld_XChk1(Fld) As String()
Const M$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
'PushI Er1, FunMsgAp_Lin(CSub, M, Fld, Ele)
Stop '
End Function

Private Function FldChk2(Fld, Ele) As String()
Const Msg1$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
PushI FldChk2, MsgAp_Lin(Msg1, Fld, Ele)
End Function

Function FldEF_XChk(F, EF As EF) As String()
If FldNm_IsStd(F) Then Exit Function
Dim Ele$
Ele = LikssDicKey(EF.E, F): If Ele = "" Then FldEF_XChk = Fld_XChk1(F): Exit Function
If EleNm_IsStd(Ele) Then Exit Function
If Not EF.E.Exists(Ele) Then FldEF_XChk = FldChk2(F, Ele)
End Function

Function Td_Str$(A As DAO.TableDef)
'See also Dbt_Stru$.  How should they
End Function

Function TdStrAy_Fny(T$()) As String()
Dim O$(), TdStr
For Each TdStr In T
    PushIAy O, TdStrFny(TdStr)
Next
TdStrAy_Fny = Ay_XWh_DistSy(O)
End Function

Function TdStrEF_XChk(TdStr, EF As EF) As String() ' Chk may return Sz=0, But Er always Sz>0
Dim Fny$(), F, O$()
Fny = TdStrFny(Stru)
For Each F In Fny
    PushIAy O, FldEF_XChk(F, EF)
Next
If Sz(0) > 0 Then
    TdStrEF_XChk = AyAp_XAdd("", O)
End If
End Function

Function TdStrFd_StrAy(A) As String()
Dim F
For Each F In TdStrFny(A)
    PushI TdStrFd_StrAy, Fd_Str(New_Fd_STD(F))
Next
End Function

Function TdStrFny(A) As String()
Dim T$, Rst$
Lin_TRstAsg A, T, Rst
If XHas_Sfx(T, "*") Then
    T = RmvSfx(T, "*")
    Rst = T & "Id " & Rst
End If
Rst = Replace(Rst, "*", T)
Rst = Replace(Rst, "|", " ")
TdStrFny = Ssl_Sy(Rst)
End Function

Function TdStrLines$(A As DAO.TableDef)
TdStrLines = JnCrLf(TdStrLy(A))
End Function

Function TdStrLy(A As DAO.TableDef) As String()
TdStrLy = Ap_Sy(TdStrLy1(A), Idxs_Ly(A.Indexes), Fds_Ly(A.Fields))
End Function

Private Function TdStrLy1$(A As DAO.TableDef)
With A
TdStrLy1 = QQ_Fmt("Tbl;?;?", .Name, TdTyStr(.Attributes))
End With
End Function

Function TdStrSk(A) As String()
Dim A1$, T$, Rst$
    A1 = XTak_Bef(A, "|")
    If A1 = "" Then Exit Function
Lin_TRstAsg A1, T, Rst
T = RmvSfx(T, "*")
Rst = Replace(Rst, "*", T)
TdStrSk = Ssl_Sy(Rst)
End Function

Private Sub ZZ()
Dim A As DAO.TableDef
Dim B$()
Dim C As Variant
Dim D As EF
TdStrLines A
TdStrLy A
TdStrSk C
End Sub

Private Sub Z()
End Sub
