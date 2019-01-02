Attribute VB_Name = "MDao_Lnk"
Option Compare Binary
Option Explicit

Function ColLnkExpFny(A$()) As String()
ColLnkExpFny = Ay_XTak_T3(A)
End Function

Sub Db_XDrp_AllLnkTbl(A As Database)
Dbtt_XDrp A, Db_XLnk_Tny(A)
End Sub

Function DbLnkSpec_XImp(A As Database, LnkSpec$()) As String()
Dim O$(), J%, T$(), L$(), W$(), U%
LnkSpecAyAsg LnkSpec, T, L, W
U = UB(LnkSpec)
For J = 0 To U
    PushAy O, Dbt_XChk_Col_ByLnkColVbl(A, T(J), L(J))
Next
If Sz(O) > 0 Then DbLnkSpec_XImp = O: Exit Function
For J = 0 To U
    Dbt_XImp_ByLnkVblCol A, T(J), L(J), W(J)
Next
DbLnkSpec_XImp = O
End Function

Function FilLin_Msg$(A$)
Dim FilNm$, Ffn$, L$
Ffn = A
FilNm = XShf_T(Ffn)
If Ffn_Exist(Ffn) Then Exit Function
FilLin_Msg = QQ_Fmt("[?] file not found [?]", FilNm, Ffn)
End Function


Function Db_LnkVblLy(A As Database) As String()
Db_LnkVblLy = AyMapPX_Sy(Db_Tny(A), "Dbt_LnkVbl", A)
End Function

Sub DrpLnkTbl()
CurDb_XDrp_AllLnkTbl
End Sub


Sub LnkSpec_XEdt()
Spnm_XEdt "Lnk"
End Sub

Sub LnkExp()
Spnm_XExp "Lnk"
End Sub

Private Property Get LnkFt$()
LnkFt = Spnm_Ft("Lnk")
End Property

Sub LnkImp()
Spnm_XImp "Lnk"
End Sub

Sub LnkSpecAyAsg(A$(), OTny$(), OLnkColVblAy$(), OWhBExprAy$())
Dim U%, J%
U = UB(A)
ReDim OTny(U)
ReDim OLnkColVblAy(U)
ReDim OWhBExprAy(U)
For J = 0 To U
    LSpecAsg A(J), OTny(J), OLnkColVblAy(J), OWhBExprAy(J)
Next
End Sub

Function LSpecLnkColVbl$(A)
Dim L$
LSpecAsg A, , L
LSpecLnkColVbl = L
End Function

Private Function NewLnkSpec(LnkSpec$) As LnkSpec
Dim Cln$():   Cln = LyCln(SplitCrLf(LnkSpec))
Dim AFx() As LnkAFil
Dim AFb() As LnkAFil
Dim ASw() As LnkASw

Dim FmFx() As LnkFmFil
Dim FmFb() As LnkFmFil
Dim FmIp() As String
Dim FmSw() As LnkFmSw
Dim FmWh() As LnkFmWh
Dim FmStu() As LnkFmStu

Dim IpFx() As LnkIpFil
Dim IpFb() As LnkIpFil
Dim IpS1() As String
Dim IpWs() As LnkIpWs

Dim StEle() As LnkStEle
Dim StExt() As LnkStExt
Dim StFld() As LnkStFld
    
    FmIp = Ssl_Sy(Ay_XWh_RmvTT(Cln, "FmIp", "|")(0))
'    FmFx = NewFmFil(Ay_XWh_RmvTT(Cln, "IpFx", "|"))
'    FmSw = NewFmSw(Ay_XWh_RmvT1(Cln, "IpSw"))
'    FmFb = NewFmFil(Ay_XWh_RmvTT(Cln, "IpFb", "|"))
'    FmWh = NewFmWh(Ay_XWh_RmvT1(Cln, "FmWh"))
    IpS1 = Ay_XWh_RmvTT(Cln, "IpS1", "|")
'    IpWs = NewIpWs(Ay_XWh_RmvTT(Cln, "IpWs", "|"))
Stop

With NewLnkSpec
    .AFx = AFx
    .AFb = AFb
    .ASw = ASw
    .FmFx = FmFx
    .FmFb = FmFb
    .FmIp = FmIp
    .FmSw = FmSw
    .FmStu = FmStu
    .FmWh = FmWh
    .IpFx = IpFx
    .IpFb = IpFb
    .IpS1 = IpS1
    .IpWs = IpWs
    .StEle = StEle
    .StExt = StExt
    .StFld = StFld
End With
End Function

Sub TTLnkFb(TT$, Fb$, Optional Fbtt)
Dbtt_XLnk_Fb CurDb, TT, Fb$, Fbtt
End Sub

Function DbInfLnkDt(A As Database) As Dt
Dim T, Dry(), C$
For Each T In Db_Tny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Set DbInfLnkDt = New_Dt("Lnk", CvNy("Tbl Connect"), Dry)
End Function

