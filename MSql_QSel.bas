Attribute VB_Name = "MSql_QSel"
Option Compare Binary
Option Explicit
Private X As New Sql_Shared

Function QSel_FF_Fm$(FF, Fm)
QSel_FF_Fm = QQSel_FF(FF) & QQFm_T(Fm)
End Function

Function QSelDis_FF_Fm$(FF, Fm)
QSelDis_FF_Fm = QQSel_FF(FF) & QQFm_T(Fm)
End Function

Function QSel_FF_Fm_Wh$(FF, Fm, WhBExpr$)
QSel_FF_Fm_Wh = QQSel_FF(FF) & QQFm_T(Fm) & QQWh_BExpr(WhBExpr)
End Function

Function QSel_FF_Fm_WhF_InAy$(FF, Fm, WhF, InAy)
Dim W$
W = FldInAy(WhF, InAy)
QSel_FF_Fm_WhF_InAy = QSel_FF_Fm_Wh(FF, Fm, W)
End Function

Function QSel_FF_ExprDic_Fm$(FF, E As Dictionary, Fm, Optional IsDis As Boolean)
'SelFFExprDicSqp = "Select" & vbCrLf & FFExprDicAsLines(FF, ExprDic)
End Function

Function QSelDis_FF_ExprDic_Fm$(FF, E As Dictionary, Fm)
QSelDis_FF_ExprDic_Fm = QSel_FF_ExprDic_Fm(FF, E, Fm, IsDis:=True)
End Function
Function QSel_FF_Into_Fm_OWh$(FF, Into, Fm, Optional WhBExpr$)
QSel_FF_Into_Fm_OWh = QQSel_FF(FF) & QQInto_T(Into) & QQFm_T(Fm) & QQWh_BExpr(WhBExpr)
End Function


Function QSel_X_Into_Fm_OWh$(X, Into, Fm, Optional WhBExpr$)
QSel_X_Into_Fm_OWh = QQSel_X(X) & QQInto_T(Into) & QQFm_T(Fm) & QQWh_BExpr(WhBExpr)
End Function

Function QSel_FF_Fm_WhFny_EqVy$(FF, Fm, Fny$(), EqVy)
QSel_FF_Fm_WhFny_EqVy = QSel_FF_Fm_Wh(FF, Fm, QQWh_FnyEqVy(Fny, EqVy))
End Function

Function QSel_FF_ExprAy_Into_Fm_OWh$(FF, Ey$(), Into, Fm, Optional WhBExpr$)
Dim X$: X = FFExprAy_AsLin(FF, Ey)
QSel_FF_ExprAy_Into_Fm_OWh = QSel_X_Into_Fm_OWh(X, Into, Fm, WhBExpr)
End Function

Function QSel_F_T$(F, T, Optional WhBExpr$)
QSel_F_T = QQ_Fmt("Select [?] from [?]?", T, F, QQWh_BExpr(WhBExpr))
End Function

Function QSel_Fm$(Fm, Optional WhBExpr$)
QSel_Fm = "Select *" & QQFm_T(Fm) & QQWh_BExpr(WhBExpr)
End Function
