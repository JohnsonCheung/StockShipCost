Attribute VB_Name = "MDao_CurDb_QQ"
Option Compare Binary
Option Explicit

Sub QQ(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
CurDb.Execute W(QQSql, Av)
End Sub

Function QQ_DTim$(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
QQ_DTim = DbqDTim(CurDb, W(QQSql, Av))
End Function

Function QQ_Rs(QQSql, ParamArray Ap()) As DAO.Recordset
Dim Av(): Av = Ap
Set QQ_Rs = Dbq_Rs(CurDb, W(QQSql, Av))
End Function

Function QQ_Tim(QQSql, ParamArray Ap()) As Date
Dim Av(): Av = Ap
QQ_Tim = Dbq_Tim(CurDb, W(QQSql, Av))
End Function

Function QQ_Val(A, ParamArray Ap())
Dim Av(): Av = Ap
QQ_Val = Dbq_Val(CurDb, W(A, Av))
End Function

Function QQ_XHas_Rec(QQSql, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
QQ_XHas_Rec = Dbq_XHas_Rec(CurDb, W(QQSql, Av))
End Function

Sub QQ_XRun(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
CurDb.Execute W(QQSql, Av)
End Sub

Private Function W$(QQSql, Av())
W = QQ_FmtAv(QQSql, Av)
End Function

Private Sub ZZ()
Dim A
Dim B()
Dim XX
QQ A, B
QQ_DTim A, B
QQ_Rs A, B
QQ_Tim A, B
QQ_Val A, B
QQ_XHas_Rec A, B
QQ_XRun A, B
End Sub

Private Sub Z()
End Sub
