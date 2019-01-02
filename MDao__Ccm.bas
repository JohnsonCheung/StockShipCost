Attribute VB_Name = "MDao__Ccm"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao__Ccm."
Private Sub Z_Db_XLnk_Ccm()
Dim CurDb As Database, IsLcl As Boolean
Set CurDb = Fb_Db(Samp_Fb_ShpRate)
IsLcl = True
GoSub Tst
Exit Sub
Tst:
    Db_XLnk_Ccm CurDb, IsLcl
    Return
End Sub

Sub Db_XLnk_Ccm(A As Database, IsLcl As Boolean)
'Ccm stands for Space-[C]ir[c]umflex-accent
'CcmTbl is ^xxx table in CurDb (pgm-database),
'          which should be same stru as N:\..._Data.accdb @ xxx
'          and   data should be copied from N:\..._Data.accdb for development purpose
'At the same time, in CurDb, there will be xxx as linked table either
'  1. In production, linking to N:\..._Data.accdb @ xxx
'  2. In development, linking to CurDb @ ^xxx
'Notes:
'  The TarFb (N:\..._Data.accdb) of each CcmTbl may be diff
'      They are stored in Description of CcmTbl manual, it is edited manually during development.
'  those xxx table in CurDb will be used in the program.
'  and ^xxx is create manually in development and should be deployed to N:\..._Data.accdb
'  assume CurDb always have some ^xxx, otherwise throw
'This Sub is to re-link the xxx in given [CurDb] to
'  1. [CurDb] if [TarFb] is not given
'  2. [TarFb] if [TarFb] is given.
Const CSub$ = CMod & "Db_XLnk_Ccm"
Dim T$()  ' All ^xxx
    T = ZCcmTny(CurDb)
    If Sz(T) = 0 Then XThw CSub, "No ^xxx table in [CurDb]", "Db", A.Name 'Assume always
ZAss CurDb, T, IsLcl ' Chk if all T after rmv ^ is in TarFb
ZLnk CurDb, T, IsLcl
End Sub
Private Sub ZAss(A As Database, CcmTny$(), IsLcl As Boolean)
Const CSub$ = CMod & "ZAss"
If Not IsLcl Then ZAss2 CurDb, CcmTny: Exit Sub ' Asserting for TarFb is stored in CcmTny's description

'Asserting for TarFb = CurDb
Dim Miss$(): Miss = ZAss1(CurDb, CcmTny)
If Sz(Miss) = 0 Then Exit Sub
XThw CSub, "[Some-missing-Tar-Tbl] in [CurDb] cannot be found according to given [CcmTny] in [CurDb]", Miss, A.Name, CcmTny, A.Name
End Sub
Private Function ZAss1(CurDb As Database, CcmTny$()) As String()
Dim N1$(): N1 = Db_Tny(CurDb)
Dim N2$(): N2 = Ay_XRmv_FstChr(CcmTny)
ZAss1 = AyMinus(N2, N1)
End Function

Private Sub ZAss2(A As Database, CcmTny$())
'Throw if any Corresponding-Table in TarFb is not found
Dim O$(), T
For Each T In CcmTny
    PushIAy O, ZAss3(A, T)
Next
Er_XHalt O
End Sub
Private Function ZAss3(A As Database, CcmTbl) As String()
Dim TarFb$
    TarFb = Dbt_Des(A, CcmTbl)
Select Case True
Case TarFb = "":             ZAss3 = MsgAp_Ly("[CcmTbl] in [CurDb] should have 'Des' which is TarFb, but this TarFb is blank", CcmTbl, CurDb.Name)
Case Ffn_NotExist(TarFb):    ZAss3 = MsgAp_Ly("[CcmTbl] in [CurDb] should have [Des] which is TarFb, but this TarFb does not exist", CcmTbl, A.Name, TarFb)
Case Not FbTbl_Exist(TarFb, XRmv_FstChr(CcmTbl)):
    ZAss3 = MsgAp_Ly("[CcmTbl] in [CurDb] should have [Des] which is TarFb, but this TarFb does not exist [Tbl-XRmv_FstChr(CcmTbl)]", CcmTbl, A.Name, TarFb, XRmv_FstChr(CcmTbl))
End Select
End Function

Private Sub ZLnk(CurDb As Database, CcmTny$(), IsLcl As Boolean)
Dim CcmTbl, TarFb$
TarFb = CurDb.Name
For Each CcmTbl In CcmTny
    If XTak_FstChr(CcmTbl) <> "^" Then Stop
    Dbtt_XLnk_Fb CurDb, XRmv_FstChr(CcmTbl), TarFb, CcmTbl
Next
End Sub
Private Function ZCcmTny(CurDb As Database) As String()
ZCcmTny = Ay_XWh_Pfx(Db_Tny(CurDb), "^")
End Function

Private Sub Z_ZCcmTny()
Dim CurDb As Database
'
Set CurDb = Fb_Db(Samp_Fb_ShpRate)
Ept = Ssl_Sy("^CurYM ^IniRate ^IniRateH ^InvD ^InvH ^YM ^YMGR ^YMGRnoIR ^YMOH ^YMRate")
GoSub Tst
Exit Sub
Tst:
    Act = ZCcmTny(CurDb)
    C
    Return
End Sub


Private Sub Z()
Z_Db_XLnk_Ccm
Z_ZCcmTny
MDao__Ccm:
End Sub
