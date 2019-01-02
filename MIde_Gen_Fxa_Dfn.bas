Attribute VB_Name = "MIde_Gen_Fxa_Dfn"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Gen_Fxa_Dfn."
Dim O$()
Property Get PjRfDfn_USR() As Dictionary
Static O As Dictionary
If IsNothing(O) Then Set O = New_Dic_LY(PjRfDfn_USR_LY, " ")
Set PjRfDfn_USR = O
End Property
Function RfNmGUIDLy(Pj_Nm) As String()
Dim RfNm
For Each RfNm In AyNz(Pj_NmRfNy_STD(Pj_Nm))
    PushI RfNmGUIDLy, RfNm & " " & RfDfn_STD(RfNm)
Next
End Function
Function Pj_NmRfNy_STD(Pj_Nm) As String()
Pj_NmRfNy_STD = Ssl_Sy(PjRfDfn_STD(Pj_Nm))
End Function
Private Property Get PjRfDfn_USR_LY() As String()
Erase O
X "MVb"
X "MIde  MVb MXls MAcs"
X "MXls  MVb"
X "MDao  MVb MDta"
X "MAdo  MVb"
X "MAdoX MVb"
X "MApp  MVb"
X "MDta  MVb"
X "MTp   MVb"
X "MSql  MVb"
X "AStkShpCst MVb MXls MAcs"
X "MAcs  MVb MXls"
PjRfDfn_USR_LY = O
Erase O
End Property

Property Get FxaNm_InDependOrder() As String()
FxaNm_InDependOrder = ASet_Sy(RelItms_DPD_ORD(NewRel(PjRfDfn_USR_LY)))
End Property

Property Get PjOwnCls() As Dictionary
Static O As Dictionary
If IsNothing(O) Then Set O = New_Dic_LY(PjOwnClsLy, " ")
Set PjOwnCls = O
End Property

Private Property Get PjOwnClsLy() As String()
Erase O
X "MVb ASet"
X "MVb WhNm"
X "MTp Blk"
X "MIde CSubBrk"
X "MIde CSubBrkMd"
X "MIde CSubBrkMth"
X "MVb DCRslt"
X "MDta Drs"
X "MDta Ds"
X "MDta Dt"
X "MVb FTIx"
X "MVb FTNo"
X "MVb FmCnt"
X "MTp Blk"
X "MTp Gp"
X "MApp LnkCol"
X "MXls LnoCnt"
X "MTp Lnx"
X "MIde Mth"
X "MIde RRCC"
X "MVb Rel"
X "MVb S1S2"
X "MSql Sql_Shared"
X "MTp SwBrk"
X "MApp New_TblImpSpec"
X "MIde VbeLoc"
X "MIde WhMd"
X "MIde WhMdMth"
X "MIde WhMth"
X "MIde WhPjMth"
PjOwnClsLy = O
Erase O
End Property

Sub PjOwnClsAss()
Const CSub$ = CMod & "PjOwnCls_VDT"
Dim ClsNy1$(), ClsNy2$()
ClsNy1 = Ay_XRmv_T1(PjOwnClsLy)
ClsNy2 = Pj_ClsNy(CurPj)
Dim X$(), Y$()
X = AyMinus(ClsNy2, ClsNy1)
Y = AyMinus(ClsNy1, ClsNy2)
If Sz(X) > 0 Or Sz(Y) > 0 Then
    Const Detail$ = "  CurPj is used to generate multi-Fxa by using the prefix of curpj as FxaNm.  In each of FxaNm should own a list of" & _
    " classes.  These list of class is defined in PjOwnClsLy.  Each of the classes should be mapped 1-and-only-1 to this definition," & _
    " otherwise it is error"
    XThw CSub, "Pj-Owning-Class-Definition error." & Detail, _
        "CurPj [Pj Owning Class Definition] [Pj Class not in def] [Def Class not in Pj]", _
        CurPj_Nm, PjOwnCls, X, Y
End If
End Sub

Property Get PjRfDfn_STD() As Dictionary
'It is a hard coded dictionary: Key is Pj_Nm, Val is RfNmss which is looking up from RfDfn_STD
Static O As Dictionary
If IsNothing(O) Then Set O = New_Dic_LY(PjRfDfn_STD_LY)
Set PjRfDfn_STD = O
End Property
Private Property Get PjRfDfn_STD_LY() As String()
Erase O
X "MVb   Scripting VBScript_RegExp_55 DAO VBIDE Office"
X "MIde  Scripting VBIDE Excel"
X "MXls  Scripting Office Excel"
X "MDao  Scripting DAO"
X "MAdo  Scripting ADODB"
X "MAdoX Scripting ADOX"
X "MApp  Scripting"
X "MDta  Scripting"
X "MTp   Scripting"
X "MSql  Scripting"
X "AStkShpCst Scripting"
X "MAcs  Scripting Office Access"
PjRfDfn_STD_LY = O
Erase O
End Property

Property Get RfDfn_STD() As Dictionary
Const CSub$ = CMod & "RfDfn_STD"
Dim O As Dictionary
Set O = New_Dic_LY(RfDfn_STD_LY)
FfnAy_XAss_Exist Ay_XRmv_3T(O.Items), CSub
Set RfDfn_STD = O
End Property

Private Property Get RfDfn_STD_LY() As String()
Erase O
X "VBA                {000204EF-0000-0000-C000-000000000046} 4  2 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
X "Access             {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07} 9  0 C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACC.OLB"
X "stdole             {00020430-0000-0000-C000-000000000046} 2  0 C:\Windows\SysWOW64\stdole2.tlb"
X "Excel              {00020813-0000-0000-C000-000000000046} 1  9 C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"
X "Scripting          {420B2830-E718-11CF-893D-00A0C9054228} 1  0 C:\Windows\SysWOW64\scrrun.dll"
X "DAO                {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28} 12 0 C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL"
X "Office             {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52} 2  8 C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL"
X "ADODB              {B691E011-1797-432E-907A-4D8C69339129} 6  1 C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
X "ADOX               {00000600-0000-0010-8000-00AA006D2EA4} 6  0 C:\Program Files (x86)\Common Files\System\ado\msadox.dll"
X "VBScript_RegExp_55 {3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5  5 C:\Windows\SysWOW64\vbscript.dll"
X "VBIDE              {0002E157-0000-0000-C000-000000000046} 5  3 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
RfDfn_STD_LY = O
Erase O
End Property

Private Sub X(A$)
PushI O, A
End Sub

Private Sub Z_FxaNm_InDependOrder()
'GoSub ZZ
GoSub T1
Exit Sub
T1:
    Ept = Ssl_Sy("MVb MXls MAdo MAdoX MApp MDta MTp MSql MDao MAcs MIde AStkShpCst")
    GoTo Tst
Tst:
    Act = FxaNm_InDependOrder
    C
    Return
ZZ:
    XClr_ImmWin
    D "Rel --------------------"
    D PjRfDfn_USR_LY
    D "Itms-DPD-ORD --------------------"
    D FxaNm_InDependOrder
    Return
End Sub

Private Sub ZZ()
PjOwnClsAss
End Sub

Private Sub Z()
Z_FxaNm_InDependOrder
End Sub
