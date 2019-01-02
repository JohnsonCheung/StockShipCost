Attribute VB_Name = "MDao_Z_Db_Dbt"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Z_Db_Dbt."

Function Dbt_Exist_Lnk(A As Database, T, S$, Cn$)
Dim I As DAO.TableDef
For Each I In A.TableDefs
    If I.Name = T Then
        If I.SourceTableName <> S Then Exit Function
        If XEns_SfxSC(I.Connect) <> XEns_SfxSC(Cn) Then Exit Function
        Dbt_Exist_Lnk = True
        Exit Function
    End If
Next
End Function

Function Dbt_Existf(A As Database, T, F) As Boolean
Ass Dbt_Exist(A, T)
Dbt_Existf = Td_XHas_FldNm(A.TableDefs(T), F)
End Function

Function Dbt_Existk(A As Database, T, K&) As Boolean
Dbt_Existk = Not Rs_IsNoRec(Dbtk_Rs(A, T, K))
End Function

Function DbtFF_Drs(A As Database, T, FF) As Drs
Set DbtFF_Drs = Dbq_Drs(A, QSel_FF_Fm(FF, T))
End Function

Function DbtFF_Dry(A As Database, T, FF) As Variant()
DbtFF_Dry = Dbq_Dry(A, QSel_FF_Fm(FF, T))
End Function

Sub DbtFF_Into_AyAp(A As Database, T, FF, ParamArray OAyAp())
Dim Drs As Drs
'Set Drs = Dbq_Drs(MthDb, "Select MchStr,MchTy,MdNm From MthMch")
Dim F, J%
For Each F In CvNy(FF)
    OAyAp(J) = DrsColInto(Drs, F, OAyAp(J))
    J = J + 1
Next
End Sub

Function DbtFF_Sql$(A As Database, T, FF)
DbtFF_Sql = QSel_FF_Fm(AyReOrdAy(Dbt_Fny(A, T), CvNy(FF)), T)
End Function

Function DbtId_Rs(A As Database, T, Id&) As DAO.Recordset
Q = QQ_Fmt("Select * From ? where ?=?", T, T, Id)
Set DbtId_Rs = A.OpenRecordset(Q)
End Function

Function DbtMsgPk$(A As Database, T)
Dim S%, K$(), O
K = Idx_Fny(DbtPIdx(A, T))
S = Sz(K)
Select Case True
Case S = 0: O = QQ_Fmt("T[?] has no Pk", T)
Case S = 1:
    If K(0) <> T & "Id" Then
        O = QQ_Fmt("T[?] has 1-field-Pk of Fld[?].  It should be [?Id]", T, K(0), T)
    End If
Case Else
    O = QQ_Fmt("T[?] has primary key.  It should have single field and name eq to table, but now it has Pk[?]", T, JnSpc(K))
End Select
DbtMsgPk = O
End Function

Function DbtPIdx(A As Database, T) As DAO.Index
Set DbtPIdx = Itr_FstItm_PrpTrue(A.TableDefs(T).Indexes, "Primary")
End Function

Function Dbt_XChk_ExpFny(A As Database, T, ExpFny0) As String()
Dim Miss$(), TFny$()
TFny = Dbt_Fny(A, T)
Miss = AyMinus(CvNy(ExpFny0), TFny)
Dbt_XChk_ExpFny = MissingFnyChk(Miss, TFny, A, T)
End Function

Function Dbt_CsvLy(A As Database, T) As String()
Dbt_CsvLy = Rs_CsvLy(Dbt_Rs(A, T))
End Function

Property Get Dbt_Des$(A As Database, T)
Dbt_Des = Dbt_Prp(A, T, C_Des)
End Property

Property Let Dbt_Des(A As Database, T, Des$)
Dbt_Prp(A, T, C_Des) = Des
End Property

Function Dbt_DftFny(A As Database, T, Optional Fny0) As String()
If IsMissing(Fny0) Then
   Dbt_DftFny = Dbt_Fny(A, T)
Else
   Dbt_DftFny = CvNy(Fny0)
End If
End Function

Function Dbt_Drs(A As Database, T) As Drs
Set Dbt_Drs = Rs_Drs(Dbt_Rs(A, T))
End Function

Function Dbt_Dry(A As Database, T) As Variant()
Dbt_Dry = Rs_Dry(Dbt_Rs(A, T))
End Function

Function Dbt_Dt(A As Database, T) As Dt
Dim Fny$(): Fny = Dbt_Fny(A, T)
Dim Dry(): Dry = Rs_Dry(Dbt_Rs(A, T))
Set Dbt_Dt = New_Dt(T, Fny, Dry)
End Function

Function Dbt_Fd_StrAy(A As Database, T) As String()
Dim F
For Each F In Dbt_Fny(A, T)
    PushI Dbt_Fd_StrAy, Dbtf_Fd_Str(A, T, F)
Next
End Function

Function Dbt_Fds(A As Database, T) As DAO.Fields
Set Dbt_Fds = A.TableDefs(T).Fields
End Function

Function Dbt_Fny(A As Database, T) As String()
Dim F As DAO.Fields
Set F = A.TableDefs(T).Fields
Dbt_Fny = Itr_Ny(F)
End Function

Function Dbt_FnyQuoted(A As Database, T) As String()
Dim O$()
O = Dbt_Fny(A, T)
If Dbt_IsXls(A, T) Then O = Ay_XQuote_SqBkt(O)
Dbt_FnyQuoted = O
End Function

Sub Dbt_FnyReSeq(A As Database, T, Fny$())
Dim F$(), J%, FF, Flds As DAO.Fields
Set Flds = A.TableDefs(T).Fields
F = AyReOrdAy(Fny, Dbt_Fny(A, T))
For Each FF In F
    J = J + 1
    Flds(FF).OrdinalPosition = J
Next
End Sub

Function Dbt_FstFldNm$(A As Database, T)
Dbt_FstFldNm = A.TableDefs(T).Fields(0).Name
End Function

Function Dbt_FstStrCol(A As DAO.Database, T) As String()
Dbt_FstStrCol = Dbtf_StrCol(A, T, Dbt_FstFldNm(A, T))
End Function

Function Dbt_Fx_LNK_TBL$(A As Database, T)
Dbt_Fx_LNK_TBL = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function Dbt_Idx(A As Database, T, Idx) As DAO.Index
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Name = Idx Then Set Dbt_Idx = I: Exit Function
Next
End Function

Function Dbt_NCol&(A As Database, T)
Dbt_NCol = A.TableDefs(T).Fields.Count
End Function

Function Dbt_NRec&(A As Database, T, Optional WhBExpr$)
Dbt_NRec = Dbq_Val(A, QQ_Fmt("Select Count(*) from [?]?", T, Wh(WhBExpr)))
End Function

Function Dbt_Pk(A As Database, T) As String()
Dbt_Pk = Idx_Fny(DbtPIdx(A, T))
End Function

Function Dbt_PkIxNm$(A As Database, T)
Obj_Nm (DbtPIdx(A, T))
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Primary Then Dbt_PkIxNm = I.Name
Next
End Function

Function Dbt_PkNm(A As Database, T)
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Primary Then Dbt_PkNm = I.Name
Next
End Function

Function Dbt_Rs(A As Database, T) As DAO.Recordset
Set Dbt_Rs = A.OpenRecordset(T)
End Function

Function Dbt_SclLy(A As Database, T) As String()
Dbt_SclLy = Td_SclLy(A.TableDefs(T))
End Function

Function Dbt_Idx_SecondaryKey(A As Database, T)
Set Dbt_Idx_SecondaryKey = Dbt_Idx(A, T, "SecondaryKey")
End Function

Function Dbt_SimTyAy(A As Database, T, Optional Fny0) As eSimTy()
Dim Fny$(): Fny = CvNy(Fny0)
Dim O() As eSimTy
   Dim U%
   ReDim O(U)
   Dim J%, F
   J = 0
   For Each F In Fny
       O(J) = DaoTy_SimTy(Dbtf_Fd(A, T, CStr(F)).Type)
       J = J + 1
   Next
Dbt_SimTyAy = O
End Function

Function Dbt_SkFny(A As Database, T) As String()
Dbt_SkFny = Idx_Fny(Dbt_SkIdx(A, T))
End Function

Function Dbt_SkValAset(A As Database, T) As ASet
Dbt_SkValAset = Idx_Fny(Dbt_SkIdx(A, T))
End Function

Function Dbt_SkAvWhStr$(A As Database, T, SkAv())
Stop '
End Function

Function Dbt_SkIdx(A As Database, T) As DAO.Index
Set Dbt_SkIdx = Idx_XAss_Sk(Dbt_Idx_SecondaryKey(A, T), T)
End Function

Function Dbt_Sq(A As Database, T, Optional InclFldNm As Boolean) As Variant()
Dbt_Sq = Rs_Sq(Dbt_Rs(A, T), InclFldNm)
End Function

Function Dbt_Sq_RESEQ(A As Database, T, ReSeqSpec$) As Variant()
Stop '
End Function

Function Dbt_Sq_RESEQ_SPEC(A As Database, T, ReSeqSpec$) As Variant()
Q = QSel_FF_Fm(ReSeqSpec_Fny(ReSeqSpec), T)
Dbt_Sq_RESEQ_SPEC = Rs_Sq(Dbq_Rs(A, Q))
End Function

Function Dbt_SrcTblNm$(A As Database, T)
Dbt_SrcTblNm = A.TableDefs(T).SourceTableName
End Function

Function Dbt_Stru$(A As Database, T)
Const CSub$ = CMod & "Dbt_Stru"
If Not Dbt_Exist(A, T) Then FunMsgAp_XDmp_Ly CSub, "[Db] has not such [Tbl]", Db_Nm(A), T: Exit Function
Dim F$()
    F = Dbt_Fny(A, T)
    If Dbt_IsXls(A, T) Then
        Dbt_Stru = T & " " & JnSpc(Ay_XQuote_SqBkt_IfNeed(F))
        Exit Function
    End If

Dim P$
    If Ay_XHas(F, T & "Id") Then
        P = " *Id"
        F = AyMinus(F, Array(T & "Id"))
    End If
Dim Sk, Rst
    Dim J%, X
    Sk = Dbt_SkFny(A, T)
    Rst = AyMinus(F, Sk)
    If Sz(Sk) > 0 Then
        For Each X In Sk
            Sk(J) = Replace(X, T, "*")
            J = J + 1
        Next
        Sk = " " & JnSpc(Ay_XQuote_SqBkt_IfNeed(Sk)) & " |"
    Else
        Sk = ""
    End If
    '
    J = 0
    For Each X In AyNz(Rst)
        Rst(J) = Replace(X, T, "*")
        J = J + 1
    Next
Rst = " " & JnSpc(Ay_XQuote_SqBkt_IfNeed(Rst))
Dbt_Stru = T & P & Sk & Rst
End Function

Function Dbt_Tim(A As Database, T) As Date
Dbt_Tim = A.TableDefs(T).Properties("LastUpdated").Value
End Function

Function Dbt_XAdd_Fd(A As Database, T, Fd As DAO.Fields) As DAO.Field2
A.TableDefs(T).Fields.Append Fd
Set Dbt_XAdd_Fd = Fd
End Function

Sub Dbt_XAdd_Fld(A As Database, T, F, Ty As DataTypeEnum, Optional Sz%, Optional Precious%)
If Dbt_Existf(A, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = DaoTy_SqlTyStr(Ty, Sz, Precious)
S = QQ_Fmt("Alter Table [?] Add Column [?] ?", T, F, Ty)
A.Execute S
End Sub

Sub Dbt_XAdd_Pfx(A As Database, T, Pfx)
Dbt_XRen A, T, Pfx & T
End Sub

Sub Dbt_XBrw(A As Database, T)
Dt_XBrw Dbt_Dt(A, T)
End Sub

Sub Dbt_XCrt_DupKeyRecTbl(A As Database, T, KeySsl$, TarTbl$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Tmp = "##" & Format(Now, "HHMMSS")
Ky = Ssl_Sy(KeySsl)
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = QQ_Fmt("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute QQ_Fmt("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, T, K)
A.Execute QQ_Fmt("Select x.* into [?] from [?] x inner join [?] a on ?", TarTbl, T, Tmp, Jn)
Dbt_XDrp A, Tmp
End Sub

Sub Dbt_XCrt_DupKeyTbl(A As Database, T, KK, TarTbl$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = Ssl_Sy(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = QQ_Fmt("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute QQ_Fmt("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, T, K)
A.Execute QQ_Fmt("Select x.* into [?] from [?] x inner join [?] a on ?", TarTbl, T, Tmp, Jn)
Dbt_XDrp A, Tmp
End Sub

Sub Dbt_XCrt_FldLisTbl(A As Database, T, TarTbl$, KyFld$, Fld$, Sep$, Optional IsMem As Boolean, Optional LisFldNm0$)
Dim LisFldNm$: LisFldNm = DftStr(LisFldNm0, Fld & "Lis")
Dim RsFm As DAO.Recordset, LasK, K, RsTo As DAO.Recordset, Lis$
A.Execute QQ_Fmt("Select ? into [?] from [?] where False", KyFld, TarTbl, T)
A.Execute QQ_Fmt("Alter Table [?] add column ? ?", TarTbl, LisFldNm, IIf(IsMem, "Memo", "Text(255)"))
Set RsFm = Dbq_Rs(A, QQ_Fmt("Select ?,? From [?] order by ?,?", KyFld, Fld, T, KyFld, Fld))
If Not Rs_IsNoRec(RsFm) Then Exit Sub
Set RsTo = Dbt_Rs(A, TarTbl)
With RsFm
    LasK = .Fields(0).Value
    Lis = .Fields(1).Value
    .MoveNext
    While Not .EOF
        K = .Fields(0).Value
        If LasK = K Then
            Lis = Lis & Sep & .Fields(1).Value
        Else
            Dr_XIns_Rs Array(LasK, Lis), RsTo
            LasK = K
            Lis = .Fields(1).Value
        End If
        .MoveNext
    Wend
End With
Dr_XIns_Rs Array(LasK, Lis), RsTo
End Sub

Sub Dbt_XCrt_Pk(A As Database, T)
Q = Tbl_CrtPkSql(T): A.Execute Q
End Sub

Sub Dbt_XCrt_Sk(A As Database, T, Fny0)
Q = QQ_Fmt("Create Unique Index SecondaryKey on ? (?)", T, JnComma(CvNy(Fny0))): A.Execute Q
End Sub

Sub Dbt_XDrp(A As Database, T)
If Dbt_Exist(A, T) Then
    A.Execute DrpTblSql(T)
End If
End Sub

Sub Dbt_XDrp_Fld(A As Database, T, Fny0)
Dim F
For Each F In AyNz(CvNy(Fny0))
    A.Execute DrpFldSql(T, F)
Next
End Sub

Sub Dbt_XImp(A As Database, T, ColLnk$())
Dbt_XDrp A, "#I" & Mid(T, 2)
Q = ColLnk_ImpSql(ColLnk, T)
A.Execute Q
End Sub

Function Dbt_LnkingFfn$(A As Database, T)
Dim Cn$
Cn = Dbt_CnStr(A, T)
If XHas_Pfx(Cn, "Excel") Or XHas_Pfx(Cn, ";DATABASE=") Then
    Dbt_LnkingFfn = XTak_BefOrAll(XTak_Aft(Cn, "DATABASE="), ";")
End If
End Function

Function Dbt_LinkingTblOrWs$(A As Database, T)
If Dbt_IsLnk(A, T) Then
    Dbt_LinkingTblOrWs = XRmv_Pfx(A.TableDefs(T).SourceTableName, "$")
End If
End Function

Sub Dbt_XReSeq_BY_FNY(A As Database, T, Fny$())
Dim TFny$(), F$(), J%, FF
TFny = Dbt_Fny(A, T)
If Sz(TFny) = Sz(Fny) Then
    F = Fny
Else
    F = Ay_XAdd_(Fny, AyMinus(TFny, Fny))
End If
For Each FF In F
    J = J + 1
    A.TableDefs(T).Fields(FF).OrdinalPosition = J
Next
End Sub

Sub Dbt_XReSeq_BY_SPEC(A As Database, T, ReSeqSpec$)
Dbt_XReSeq_BY_FNY A, T, ReSeqSpec_Fny(ReSeqSpec)
End Sub

Sub Dbt_XReStru(A As Database, Stru, EF As EF)
Dim DrpFld$(), NewFld$(), T$, FnyO$(), FnyN$()
T = Lin_T1(Stru)
FnyO = Dbt_Fny(A, T)
FnyN = TdStrFny(Stru)

DrpFld = AyMinus(FnyO, FnyN)
NewFld = AyMinus(FnyN, FnyO)
Dbt_XDrp_Fld A, T, DrpFld
DbtFFAdd A, T, NewFld, EF
End Sub

Sub Dbt_XRen(A As Database, T, ToTbl)
Fb_Db(A.Name).TableDefs(T).Name = ToTbl
End Sub

Sub Dbt_XRen_Col(A As Database, T, Fm, NewCol)
Fb_Db(A.Name).TableDefs(T).Fields(Fm).Name = NewCol
End Sub

Sub Dbt_XSet_FDes(A As Database, T, FDes As Dictionary)
Stop '
'Dbtf_Des(A, T, B.F) = B.Des
End Sub

Sub Dbt_XUpd_IdFld(A As Database, T, Fld)
Dim D As New Dictionary, J&, Rs As DAO.Recordset, Id$, V
Id = Fld & "Id"
Set Rs = Dbq_Rs(A, QQ_Fmt("Select ?,?Id from [?]", Fld, Fld, T))
With Rs
    While Not .EOF
        .Edit
        V = .Fields(0).Value
        If D.Exists(V) Then
            .Fields(1).Value = D(V)
        Else
            .Fields(1).Value = J
            D.Add V, J
            J = J + 1
        End If
        .Update
        .MoveNext
    Wend
End With
End Sub

Sub Dbt_XUpd_Seq(A As Database, T, SeqFldNm$, Optional RestFny0, Optional IncFny0)
'Assume T is sorted
'
'Update A->T->SeqFldNm using RestFny0,IncFny0, assume the table has been sorted
'Update A->T->SeqFldNm using OrdFny0, RestFny0,IncFny0
Dim RestFny$(), IncFny$(), Sql$
Dim LasRestVy(), LasIncVy(), Seq&, OrdS$, Rs As DAO.Recordset
'OrdFny RestAy IncAy Sql
RestFny = CvNy(RestFny0)
IncFny = CvNy(IncFny0)
If Sz(RestFny) = 0 And Sz(IncFny) = 0 Then
    With A.OpenRecordset(T)
        Seq = 1
        While Not .EOF
            .Edit
            .Fields(SeqFldNm) = Seq
            Seq = Seq + 1
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Exit Sub
End If
'--
Set Rs = A.OpenRecordset(T) ', RecordOpenOptionsEnum.dbOpenForwardOnly, dbForwardOnly)
With Rs
    While Not .EOF
        If Rs_IsBrk(Rs, RestFny, LasRestVy) Then
            Seq = 1
            LasRestVy = Rs_Dr_ByFF(Rs, RestFny)
            LasIncVy = Rs_Dr_ByFF(Rs, IncFny)
        Else
            If Rs_IsBrk(Rs, IncFny, LasIncVy) Then
                Seq = Seq + 1
                LasIncVy = Rs_Dr_ByFF(Rs, IncFny)
            End If
        End If
        .Edit
        .Fields(SeqFldNm).Value = Seq
        .Update
        .MoveNext
    Wend
End With
End Sub

Sub Dbt_XUpd_ToDteFld(A As Database, T, ToDteFld$, KeyFld$, FmDteFld$)
Dim ToDte() As Date, J&
ToDte = Dbt_XUpd_ToDteFld__1(A, T, KeyFld, FmDteFld)
With Dbt_Rs(A, T)
    While Not .EOF
        .Edit
        .Fields(ToDteFld).Value = ToDte(J): J = J + 1
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

Function Dbt_XUpd_ToDteFld__1(A As Database, T, KeyFld$, FmDteFld$) As Date()
Dim K$(), FmDte() As Date, ToDte() As Date, J&, CurKey$, NxtKey$, NxtFmDte As Date
With Dbt_Rs(A, T)
    While Not .EOF
        Push FmDte, .Fields(FmDteFld).Value
        Push K, .Fields(KeyFld).Value
        .MoveNext
    Wend
End With
Dim U&
U = UB(K)
ReDim ToDte(U)
For J = 0 To U - 1
    CurKey = K(J)
    NxtKey = K(J + 1)
    NxtFmDte = FmDte(J + 1)
    If CurKey = NxtKey Then
        ToDte(J) = DateAdd("D", -1, NxtFmDte)
    Else
        ToDte(J) = DateSerial(2099, 12, 31)
    End If
Next
ToDte(U) = DateSerial(2099, 12, 31)
Dbt_XUpd_ToDteFld__1 = ToDte
End Function

Function Dbtf_AyInto(A As DAO.Database, T, F, OInto)
Dim Rs As DAO.Recordset
Set Rs = Dbtf_Rs(A, T, F)
Dbtf_AyInto = Rs_AyInto(Rs, OInto)
End Function

Function Dbtf_EleScl$(A As DAO.Database, T, F)
Dim Td As DAO.TableDef, Fd As DAO.Field2
Set Td = A.TableDefs(T)
Set Fd = Td.Fields(F)
Dbtf_EleScl = FdEleScl(Fd)
End Function

Function Dbtf_Fd(A As Database, T, F) As DAO.Field2
Set Dbtf_Fd = A.TableDefs(T).Fields(F)
End Function

Function Dbtf_FdInfDr(A As Database, T, F) As Variant()
Dim FF  As DAO.Field
Set FF = A.TableDefs(T).Fields(F)
With FF
    Dbtf_FdInfDr = Array(F, IIf(Dbtf_IsPk(A, T, F), "*", ""), DaoShtTyStr_DaoTy(.Type), .Size, .DefaultValue, .Required, Fd_Des(FF))
End With
End Function

Function Dbtf_Fd_Str$(A As Database, T, F)
Dbtf_Fd_Str = Fd_Str(Dbtf_Fd(A, T, F))
End Function

Function Dbtf_IntAy(A As DAO.Database, T, F) As Integer()
Q = QQ_Fmt("Select [?] from [?]", F, T)
Dbtf_IntAy = Dbq_IntAy(A, Q)
End Function

Function Dbtf_IsPk(A As Database, T, F) As Boolean
Dbtf_IsPk = Ay_XHas(Dbt_Pk(A, T), F)
End Function

Function Dbtf_NxtId&(A As Database, T, Optional F)
Dim S$: S = QQ_Fmt("select Max(?) from ?", Dft(F, T), T)
Dbtf_NxtId = Dbq_Val(A, S) + 1
End Function

Function Dbtf_Rs(A As DAO.Database, T, F) As DAO.Recordset
Set Dbtf_Rs = Dbq_Rs(A, SelTFSql(T, F))
End Function

Function Dbtf_StrCol(A As DAO.Database, T, F) As String()
Dbtf_StrCol = Dbtf_AyInto(A, T, F, Dbtf_StrCol)
End Function

Function Dbtf_DaoTy(A As Database, T, F) As DAO.DataTypeEnum
Dbtf_DaoTy = A.TableDefs(T).Fields(F).Type
End Function

Function Dbtf_DaoShtTyStr$(A As Database, T, F)
Dbtf_DaoShtTyStr = DaoTy_DaoShtTyStr(Dbtf_DaoTy(A, T, F))
End Function

Function Dbtf_Val(A As Database, T, F)
Dbtf_Val = A.TableDefs(T).OpenRecordset.Fields(F).Value
End Function

Sub Dbtf_XAdd_Expr(A As Database, T, F, Expr$, Optional Ty As DAO.DataTypeEnum = dbText)
A.TableDefs(T).Fields.Append New_Fd(F, Ty, Expr:=Expr)
End Sub

Sub Dbtf_XChg_DteToTxt(A As Database, T, F)
A.Execute QQ_Fmt("Alter Table [?] add column [###] text(12)", T)
A.Execute QQ_Fmt("Update [?] set [###] = Format([?],'YYYY-MM-DD')", T, F)
A.Execute QQ_Fmt("Alter Table [?] Drop Column [?]", T, F)
A.Execute QQ_Fmt("Alter Table [?] Add Column [?] text(12)", T, F)
A.Execute QQ_Fmt("Update [?] set [?] = [###]", T, F)
A.Execute QQ_Fmt("Alter Table [?] Drop Column [###]", T)
End Sub

Sub DbtVy_XIns(A As Database, T, Vy)
Dr_XIns_Rs CvVV(Vy), Dbt_Rs(A, T)
End Sub

Function Dbtk_Rs(A As Database, T, K&) As DAO.Recordset
Set Dbtk_Rs = A.OpenRecordset(QSel_Fm(T, WhK(K, T)))
End Function

Function Dbtkf_Val(A As Database, T, K&, F) ' K is Pk value
Dbtkf_Val = Dbq_Val(A, QSel_FF_Fm_Wh(T, F, WhK(K, T)))
End Function

Sub Dbtt_Brw(A As Database, TT0)
Ay_XDoPX CvNy(TT0), "Dbt_XBrw", A
End Sub

Sub Dbtt_Imp(A As Database, TT)
Dim Tny$(), J%, S$
Tny = CvNy(TT)
For J = 0 To UB(Tny)
    Dbt_XDrp A, "#I" & Tny(J)
    S = QQ_Fmt("Select * into [#I?] from [?]", Tny(J), Tny(J))
    A.Execute S
Next
End Sub

Function Dbtt_XLnk_Fb(A As Database, TT$, Fb$, Optional Fb_Tny0) As String()
Dim Tny$(), Fb_Tny$(), J%, T
Tny = CvNy(TT)
Fb_Tny = CvNy(Fb_Tny0)
    Select Case Sz(Fb_Tny)
    Case Sz(Tny)
    Case 0:    Fb_Tny = Tny
    Case Else: XThw CSub, "[TT]-[Sz1] and [Fbtt]-[Sz2] are diff.  (@Dbtt_XLnk_Fb)", TT, Sz(Tny), Fb_Tny, Sz(Fb_Tny)
    End Select
Dim Cn$: Cn = Fb_CnStr_DAO(Fb)
For Each T In Tny
    PushIAy Dbtt_XLnk_Fb, Dbt_XLnk(A, CStr(T), Fb_Tny(J), Cn)
    J = J + 1
Next
End Function

Function Dbtt_Stru(A As Database, TT)
Dim T, O$()
For Each T In Ay_XSrt(CvNy(TT))
    PushI O, Dbt_Stru(A, CStr(T))
Next
Dbtt_Stru = JnCrLf(Ay_XAlign_1T(O))
End Function

Sub Dbtt_XCrtPk(A As Database, TT0)
Ay_XDoPX CvNy(TT0), "Dbt_XCrt_Pk", A
End Sub

Sub Dbtt_XDrp(A As Database, TT)
Dim T
For Each T In CvNy(TT)
    Dbt_XDrp A, T
Next
End Sub

Sub Dbtt_XRen_ByXAdd_Pfx(A As Database, TT0, Pfx)
Ay_XDoAXB CvNy(TT0), "Dbt_XAdd_Pfx", A, Pfx
End Sub

Function CvVV(VV) As Variant()
Select Case True
Case IsVy(VV): CvVV = VV
Case IsArray(VV)
    Dim I
    For Each I In AyNz(VV)
        PushI CvVV, I
    Next
Case Else: CvVV = Array(VV)
End Select
End Function

Function Idx_XAss_Sk(A As DAO.Index, T, Optional Fun$ = "Idx_XAss_Sk") As DAO.Index 'SkIdx is DaoIdx with name as 'SecondaryKey'
Const CSub$ = CMod & "SkIdx_XAss"
If IsNothing(A) Then Exit Function
If Not A.Unique Then XThw CSub, "Idx-SecondaryKey should be Unique", "Tbl", T
If A.Primary Then XThw CSub, "Idx-SecondaryKey cannot have a flag PrimaryKey=True", "Tbl", T
Set Idx_XAss_Sk = A
End Function

Function TF_Fd_Str$(T, F)
TF_Fd_Str = Dbtf_Fd_Str(CurDb, T, F)
End Function

Private Sub ZZ_DbtPutTp()

End Sub

Private Sub ZZ_DbtSpec_XReSeq()
Dbt_XReSeq_BY_SPEC CurDb, "ZZ_Dbt_XUpd_Seq", "Permit PermitD"
End Sub

Private Sub ZZ_Dbt_XCrt_DupKeyRecTbl()
Tbl_XDrp "#A #B"
DoCmd.RunSQL "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_Dbt_XUpd_Seq"
Dbt_XCrt_DupKeyRecTbl CurDb, "#A", "Sku BchNo", "#B"
Tbl_XBrw "#B"
Stop
Tbl_XDrp "#B"
End Sub

Private Sub ZZ_Dbt_XCrt_DupKeyTbl()
TT_XDrp "#A #B"
DoCmd.RunSQL "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_Dbt_XUpd_Seq"
Dbt_XCrt_DupKeyTbl CurDb, "#A", "Sku BchNo", "#B"
TT_XBrw "#B"
Stop
TT_XDrp "#B"
End Sub

Private Sub ZZ_Dbt_XReSeq_BY_SPEC()
Dbt_XReSeq_BY_SPEC CurDb, "ZZ_Dbt_XUpd_Seq", "Permit PermitD"
End Sub

Private Sub ZZ_Dbt_XUpd_Seq()
DoCmd.SetWarnings False
DoCmd.RunSQL "Select * into [#A] from ZZ_Dbt_XUpd_Seq order by Sku,PermitDate"
DoCmd.RunSQL "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
Dbt_XUpd_Seq CurDb, "#A", "BchRateSeq", "Sku", "Sku Rate"
Tbl_XOpn "#A"
Stop
DoCmd.RunSQL "Drop Table [#A]"
End Sub

Private Sub ZZ_Dbt_XUpd_ToDteFld()
DoCmd.RunSQL "Select * into [#A] from ZZ_Dbt_XUpd_ToDteFld order by Sku,PermitDate"
Dbt_XUpd_ToDteFld CurDb, "#A", "PermitDateEnd", "Sku", "PermitDate"
Stop
Tbl_XDrp "#A"
End Sub

Private Sub Z_Dbt_Pk()
ZZ:
    Dim A As Database
    Set A = Fb_Db(Samp_Fb_Duty_Dta)
    Dim Dr(), Dry(), T
    For Each T In Db_Tny(A)
        Erase Dr
        Push Dr, T
        PushIAy Dr, Dbt_Pk(A, T)
        PushI Dry, Dr
    Next
    Dry_XBrw Dry
    Exit Sub
End Sub

Private Sub ZZ()
Dim A As Database
Dim B
Dim C As DAO.Fields
Dim D As DataTypeEnum
Dim E%
Dim F$
Dim G As Boolean
Dim H()
Dim I$()
Dim J&
Dim K As EF
Dim L As Dictionary
Dim M As DAO.Index
Dim O As DAO.Database
Dim XX
Dbt_Exist_Lnk A, B, F, F
Dbt_Existf A, B, B
Dbt_Existk A, B, J
DbtFF_Drs A, B, B
DbtFF_Dry A, B, B
DbtFF_Into_AyAp A, B, B, H
DbtFF_Sql A, B, B
DbtId_Rs A, B, J
DbtMsgPk A, B
DbtPIdx A, B
Dbt_XChk_ExpFny A, B, B
Dbt_CsvLy A, B
Dbt_DftFny A, B, B
Dbt_Drs A, B
Dbt_Dry A, B
Dbt_Dt A, B
Dbt_Fd_StrAy A, B
Dbt_Fds A, B
Dbt_Fny A, B
Dbt_FnyQuoted A, B
Dbt_FnyReSeq A, B, I
Dbt_FstFldNm A, B
Dbt_FstStrCol O, B
Dbt_Fx_LNK_TBL A, B
Dbt_Idx A, B, B
Dbt_NCol A, B
Dbt_NRec A, B, F
Dbt_Pk A, B
Dbt_PkIxNm A, B
Dbt_PkNm A, B
Dbt_Rs A, B
Dbt_SclLy A, B
Dbt_SimTyAy A, B, B
Dbt_SkAvWhStr A, B, H
Dbt_SkIdx A, B
Dbt_Sq A, B, G
Dbt_Sq_RESEQ A, B, F
Dbt_Sq_RESEQ_SPEC A, B, F
Dbt_SrcTblNm A, B
Dbt_Stru A, B
Dbt_Tim A, B
Dbt_XAdd_Fd A, B, C
Dbt_XAdd_Fld A, B, B, D, E, E
Dbt_XAdd_Pfx A, B, B
Dbt_XBrw A, B
Dbt_XCrt_DupKeyRecTbl A, B, F, F
Dbt_XCrt_DupKeyTbl A, B, B, F
Dbt_XCrt_FldLisTbl A, B, F, F, F, F, G, F
Dbt_XCrt_Pk A, B
Dbt_XCrt_Sk A, B, B
Dbt_XDrp A, B
Dbt_XDrp_Fld A, B, B
Dbt_XImp A, B, I
Dbt_LnkingFfn A, B
Dbt_LinkingTblOrWs A, B
Dbt_XReSeq_BY_FNY A, B, I
Dbt_XReSeq_BY_SPEC A, B, F
Dbt_XReStru A, B, K
Dbt_XRen A, B, B
Dbt_XRen_Col A, B, B, B
Dbt_XSet_FDes A, B, L
Dbt_XUpd_IdFld A, B, B
Dbt_XUpd_Seq A, B, F, B, B
Dbt_XUpd_ToDteFld A, B, F, F, F
Dbt_XUpd_ToDteFld__1 A, B, F, F
Dbtf_AyInto O, B, B, B
Dbtf_EleScl O, B, B
Dbtf_Fd A, B, B
Dbtf_FdInfDr A, B, B
Dbtf_Fd_Str A, B, B
Dbtf_IntAy O, B, B
Dbtf_IsPk A, B, B
Dbtf_NxtId A, B, B
Dbtf_Rs O, B, B
Dbtf_StrCol O, B, B
Dbtf_Val A, B, B
Dbtf_XChg_DteToTxt A, B, B
Dbtk_Rs A, B, J
Dbtkf_Val A, B, J, B
Dbtt_Brw A, B
Dbtt_Imp A, B
Dbtt_XLnk_Fb A, F, F, B
Dbtt_Stru A, B
Dbtt_XCrtPk A, B
Dbtt_XDrp A, B
Dbtt_XRen_ByXAdd_Pfx A, B, B
TF_Fd_Str B, B
XX = Dbt_Des(A, B)
End Sub

Private Sub Z()
Z_Dbt_Pk
End Sub
