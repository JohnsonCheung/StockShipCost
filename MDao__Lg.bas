Attribute VB_Name = "MDao__Lg"
Option Compare Binary
Option Explicit
Private XSchm$()
Private X_W As Database
Private X_L As Database
Private X_Sess&
Private X_Msg&
Private X_Lg&
Private O$() ' Used by Pth_EntAyR

Sub CurLgLis(Optional Sep$ = " ", Optional Top% = 50)
D CurLgLy(Sep, Top)
End Sub

Function CurLgLy(Optional Sep$ = " ", Optional Top% = 50) As String()
CurLgLy = Rs_Ly(CurLgRs(Top), Sep)
End Function
Private Function Rs_Ly(A As DAO.Database, Sep$) As String()

End Function
Function CurLgRs(Optional Top% = 50) As DAO.Recordset
Set CurLgRs = L.OpenRecordset(QQ_Fmt("Select Top ? x.*,Fun,MsgTxt from Lg x left join Msg a on x.Msg=a.Msg order by Sess desc,Lg", Top))
End Function

Sub CurSessLis(Optional Sep$ = " ", Optional Top% = 50)
D CurSessLy(Sep, Top)
End Sub

Function CurSessLy(Optional Sep$, Optional Top% = 50) As String()
CurSessLy = Rs_Ly(CurSessRs(Top), Sep)
End Function

Function CurSessRs(Optional Top% = 50) As DAO.Recordset
Set CurSessRs = L.OpenRecordset(QQ_Fmt("Select Top ? * from sess order by Sess desc", Top))
End Function
Private Function Dbq_Val(A As Database, Q)
Stop
End Function
Private Function CvSess&(A&)
If A > 0 Then CvSess = A: Exit Function
CvSess = Dbq_Val(L, "select Max(Sess) from Sess")
End Function
Private Sub XEnsMsg(Fun$, MsgTxt$)
With L.TableDefs("Msg").OpenRecordset
    .Index = "Msg"
    .Seek "=", Fun, MsgTxt
    If .NoMatch Then
        .AddNew
        !Fun = Fun
        !MsgTxt = MsgTxt
        X_Msg = !Msg
        .Update
    Else
        X_Msg = !Msg
    End If
End With
End Sub


Private Sub XEns_Sess()
If X_Sess > 0 Then Exit Sub
With L.TableDefs("Sess").OpenRecordset
    .AddNew
    X_Sess = !Sess
    .Update
    .Close
End With
End Sub

Private Property Get L() As Database
On Error GoTo X
If IsNothing(X_L) Then
    LgOpn
End If
Set L = X_L
Exit Property
X:
Dim E$, ErNo%
ErNo = Err.Number
E = Err.Description
If ErNo = 3024 Then
    'LgSchmImp
    LgCrt_v1
    LgOpn
    Set L = X_L
    Exit Property
End If
FunMsgNyAp_XDmp CSub, "Cannot open LgDb", "Er ErNo", E, ErNo
Stop
End Property

Sub Lg(Fun$, MsgTxt$, ParamArray Ap())
XEns_Sess
XEnsMsg Fun, MsgTxt
WrtLg Fun, MsgTxt
Dim Av(): Av = Ap
If Sz(Av) = 0 Then Exit Sub
Dim J%, V
With L.TableDefs("LgV").OpenRecordset
    For Each V In Av
        .AddNew
        !Lines = Var_Lines(V)
        .Update
    Next
    .Close
End With
End Sub

Private Sub Rs_Asg(A As DAO.Recordset, ParamArray OAp())

End Sub

Sub LgAsg(A&, OSess&, ODTim$, OFun$, OMsgTxt$)
Dim Q$
Q = QQ_Fmt("select Fun,MsgTxt,Sess,x.CrtTim from Lg x inner join Msg a on x.Msg=a.Msg where Lg=?", A)
Dim D As Date
Rs_Asg L.OpenRecordset(Q), OFun, OMsgTxt, OSess, D
ODTim = Dte_DTim(D)
End Sub

Sub LgBeg()
Lg ".", "Beg"
End Sub

Sub LgBrw()
Ft_XBrw LgFt
End Sub

Sub LgCls()
On Error GoTo Er
X_L.Close
Er:
Set X_L = Nothing
End Sub
Private Sub Fb_XCrt(A)
Stop '
End Sub
Private Function Fb_Db(A) As Database
Stop '
End Function
Private Sub TdAddId(A As DAO.TableDef)
End Sub
Sub LgCrt()
'Fb_XCrt LgFb
'Dim Db As Database, T As DAO.TableDef
'Set Db = Fb_Db(LgFb)
''
'Set T = New DAO.TableDef
'T.Name = "Sess"
'TdAddId T
'Td_XAdd_TimStampFld T, "Dte"
'Db.TableDefs.Append T
''
'Set T = New DAO.TableDef
'T.Name = "Msg"
'TdAddId T
'TdAddTxtFld T, "Fun"
'TdAddTxtFld T, "MsgTxt"
'Td_XAdd_TimStampFld T, "Dte"
'Db.TableDefs.Append T
''
'Set T = New DAO.TableDef
'T.Name = "Lg"
'TdAddId T
'Td_XAdd_LngFld T, "Sess"
'Td_XAdd_LngFld T, "Msg"
'Td_XAdd_TimStampFld T, "Dte"
'Db.TableDefs.Append T
''
'Set T = New DAO.TableDef
'T.Name = "LgV"
'TdAddId T
'Td_XAdd_LngFld T, "Lg"
'Td_XAdd_LngTxt T, "Val"
'Db.TableDefs.Append T
'
'Dbtt_XCrtPk Db, "Sess Msg Lg LgV"
'Dbt_XCrt_Sk Db, "Msg", "Fun MsgTxt"
End Sub

Sub LgCrt_v1()
Dim Fb$
Fb = LgFb
If Ffn_Exist(Fb) Then Exit Sub
'DbCrtSchm Fb_XCrt(Fb), LgSchmLines
End Sub

Property Get LgDb() As Database
Set LgDb = L
End Property

Sub LgDb_XBrw()
'Acs.OpenCurrentDatabase LgFb
'AcsVis Acs
End Sub

Sub LgEnd()
Lg ".", "End"
End Sub

Property Get LgFb$()
LgFb = LgPth & LgFn
End Property

Property Get LgFn$()
LgFn = "Lg.accdb"
End Property

Property Get LgFt$()
Stop '
End Property
Private Sub X(A$)
PushI XSchm, A
End Sub
Property Get LgSchm() As String()
If Sz(XSchm) = 0 Then
X "E Mem | Mem Req AlwZLen"
X "E Txt | Txt Req"
X "E Crt | Dte Req Dft=Now"
X "E Dte | Dte"
X "E Amt | Cur"
X "F Amt * | *Amt"
X "F Crt * | CrtDte"
X "F Dte * | *Dte"
X "F Txt * | Fun * Txt"
X "F Mem * | Lines"
X "T Sess | * CrtDte"
X "T Msg  | * Fun *Txt | CrtDte"
X "T Lg   | * Sess Msg CrtDte"
X "T LgV  | * Lg Lines"
X "D . Fun | Function name that call the log"
X "D . Fun | Function name that call the log"
X "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt"
X "D . Msg | ..."
End If
LgSchm = XSchm
End Property

Sub LgKill()
LgCls
If Ffn_Exist(LgFb) Then Kill LgFb: Exit Sub
Debug.Print "LgFb-[" & LgFb & "] not exist"
End Sub

Function LgLinesAy(A&) As Variant()
Dim Q$
Q = QQ_Fmt("Select Lines from LgV where Lg = ? order by LgV", A)
'LgLinesAy = Rs_Ay(L.OpenRecordset(Q))
End Function

Sub LgLis(Optional Sep$ = " ", Optional Top% = 50)
CurLgLis Sep, Top
End Sub

Function LgLy(A&) As String()
Dim Fun$, MsgTxt$, DTim$, Sess&, Sfx$
LgAsg A, Sess, DTim, Fun, MsgTxt
Sfx = QQ_Fmt(" @? Sess(?) Lg(?)", DTim, Sess, A)
'LgLy = FunDotMsg_Fmt(Fun & Sfx, MsgTxt, LgLinesAy(A))
Stop '
End Function

Private Sub LgOpn()
Set X_L = Fb_Db(LgFb)
End Sub

Property Get LgPth$()
Static Y$
'If Y = "" Then Y = PgmPth & "Log\": Pth_XEns Y
LgPth = Y
End Property

Sub FmCntDmp(A As FmCnt)
Debug.Print FmCntStr(A)
End Sub

Sub SessBrw(Optional A&)
Ay_XBrw SessLy(CvSess(A))
End Sub

Function SessLgAy(A&) As Long()
Dim Q$
Q = QQ_Fmt("select Lg from Lg where Sess=? order by Lg", A)
'SessLgAy = Dbq_LngAy(L, Q)
End Function

Sub SessLis(Optional Sep$ = " ", Optional Top% = 50)
CurSessLis Sep, Top
End Sub

Function SessLy(Optional A&) As String()
Dim LgAy&()
LgAy = SessLgAy(A)
SessLy = AyOfAy_Ay(AyMap(LgAy, "LgLy"))
End Function

Function SessNLg%(A&)
SessNLg = Dbq_Val(L, "Select Count(*) from Lg where Sess=" & A)
End Function

Private Sub WrtLg(Fun$, MsgTxt$)
With L.TableDefs("Lg").OpenRecordset
    .AddNew
    !Sess = X_Sess
    !Msg = X_Msg
    X_Lg = !Lg
    .Update
End With
End Sub

Private Sub Z_Lg()
LgKill
Debug.Assert Dir(LgFb) = ""
LgBeg
Debug.Assert Dir(LgFb) = LgFn
End Sub


Private Sub Z()
Z_Lg
MDao_Lg:
End Sub
