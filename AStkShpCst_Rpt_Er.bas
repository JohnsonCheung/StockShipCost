Attribute VB_Name = "AStkShpCst_Rpt_Er"
Option Compare Binary
Option Explicit

Sub XRptEr()
Er_XHalt Ap_Sy(MB52_8601_8701_Missing)
End Sub

Private Property Get MB52_8601_8701_Missing() As String()
Const M$ = "Column-[Plant] must have value 8601 or 8701"
Const Wh$ = "Plant in ('8601','8701')"
Dim Fx$
Fx = IFxMB52
Dbt_XLnk_Fx W, "#A", Fx
If Dbt_NRec(W, "#A", Wh) = 0 Then
    MB52_8601_8701_Missing = _
        FunMsgNyAp_Ly(CSub, M, "MB52-File Worksheet", Fx, "Sheet1")
End If
WDrp "#A"
End Property

