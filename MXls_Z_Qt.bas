Attribute VB_Name = "MXls_Z_Qt"
Option Compare Binary
Option Explicit
Function Qt_FbtStr$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
Qt_FbtStr = QQ_Fmt("[?].[?]", CnStr_DtaSrcVal(CnStr), Tbl)
End Function
