Attribute VB_Name = "MDao_Att_Inf"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Att_Inf."

Function AttFfn$(A)
'Return Fst-Ffn-of-Att-A
AttFfn = RsMovFst(New_AttRs(A).AttRs)!FileName
End Function

Function AttFilCnt%(A)
AttFilCnt = DbAtt_NFil(CurDb, A)
End Function

Function Att_FnAy(A) As String()
Att_FnAy = DbAtt_FnAy(CurDb, "AA")
End Function

Property Get AttFny() As String()
AttFny = Itr_Ny(Db_FstAttRs(CurDb).AttRs.Fields)
End Property

Function Att_FstFn$(A)
Att_FstFn = DbAtt_FstFn(CurDb, A)
End Function

Function Att_XHas_OnlyOneFile(A$) As Boolean
Att_XHas_OnlyOneFile = DbAtt_XHas_OnlyOneFile(CurDb, A)
End Function

Function Att_IsOld(A$, Ffn$) As Boolean
Const CSub$ = CMod & "Att_IsOld"
Dim TAtt As Date, TFfn As Date, AttIs$
TAtt = Att_Tim(A)
TFfn = Ffn_Tim(Ffn)
AttIs = IIf(TAtt > TFfn, "new", "old")
Dim M$
M = "Att is " & AttIs
XDmp_Ly CSub, M, "Att Ffn AttTim Ffn_Tim AttIs-Old-or-New?", A, Ffn, TAtt, TFfn, AttIs
End Function

Property Get AttNy() As String()
AttNy = CurDb_AttNy
End Property

Function Att_Sz&(A)
Att_Sz = TF_Val_BySkVal("Att", "FilSz", A)
End Function

Function Att_Tim(A) As Date
Att_Tim = TF_Val_BySkVal("Att", "FilTim", A)
End Function

Property Get CurDb_AttNy() As String()
CurDb_AttNy = Db_AttNy(CurDb)
End Property

Function DbAtt_NFil%(A As Database, Att)
'DbAtt_NFil = DbAtt_AttRs(A, Att).AttRs.RecordCount
DbAtt_NFil = AttRs_NFil(DbAtt_AttRs(A, Att))
End Function

Function DbAtt_FnAy(A As Database, Att$) As String()
Dim T As DAO.Recordset ' AttTbl_Rs
Dim F As DAO.Recordset ' AttFldRs
Set T = DbAttTbl_Rs(A, Att)
If T.EOF And T.BOF Then Exit Function
Set F = T.Fields("Att").Value
DbAtt_FnAy = Rs_Sy(F, "FileName")
End Function

Function DbAtt_FstFn(A As Database, Att)
DbAtt_FstFn = AttRs_FstFn(DbAtt_AttRs(A, Att))
End Function

Function DbAtt_XHas_OnlyOneFile(A As Database, Att$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & DbAtt_AttRs(A, Att).AttRs.RecordCount
DbAtt_XHas_OnlyOneFile = DbAtt_AttRs(A, Att).AttRs.RecordCount = 1
End Function

Function Db_AttNy(A As Database) As String()
Q = "Select AttNm from Att order by AttNm": Db_AttNy = Rs_Sy(A.OpenRecordset(Q))
End Function

Function DbAttTbl_Rs(A As Database, AttNm$) As DAO.Recordset
Set DbAttTbl_Rs = A.OpenRecordset(QQ_Fmt("Select * from Att where AttNm='?'", AttNm))
End Function

Private Sub Z_Att_FnAy()
Fb_CurDb Samp_Fb_ShpRate
D Att_FnAy("AA")
CurDb_XCls
End Sub

Private Sub Z()
Z_Att_FnAy
MDao_Att_Inf:
End Sub
