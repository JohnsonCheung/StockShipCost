Attribute VB_Name = "MSql_QDel"
Option Compare Binary
Option Explicit

Function QDlt_Fm_OWh$(T, Optional BExpr$)
QDlt_Fm_OWh = "Delete * from [" & T & "]" & WhBExprSqp(BExpr)
End Function

Function QDlt_Fm_F_InASet(T, F, S As ASet, Optional SqlWdt% = 3000) As String()
Dim Hdr$
    Hdr = QQ_Fmt("Delete * from [T] Where ", T)
Dim Ey$()
    Ey = ASet_BInExprAy(S, F, SqlWdt - Len(Hdr))
Dim E
For Each E In Ey
    PushI QDlt_Fm_F_InASet, Hdr & E
Next
End Function
Function ASet_SqlInLisAy(A As ASet, Wdt%) As String()
Dim Q$
    Q = Val_SqlQChr(ASet_FstItm(A))

End Function
Function Val_SqlQChr$(A)

End Function

Function ASet_BInExprAy(A As ASet, F, ExprWdt%) As String()
If ASet_IsEmp(A) Then XThw CSub, "Given ASet is empty", "F ExprWdt", F, ExprWdt
Dim Pfx$
    Pfx = "[" & F & "] in "
Dim InLis
For Each InLis In ASet_SqlInLisAy(A, ExprWdt)
    PushI ASet_BInExprAy, Pfx & InLis
Next
End Function
