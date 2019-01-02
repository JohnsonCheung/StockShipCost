Attribute VB_Name = "MAdo_Fb_Fbq"
Option Compare Binary
Option Explicit
Const Samp_Fb_DutyPrepare$ = ""
Function FbqAdoDrs(A$, Q$) As Drs
Set FbqAdoDrs = ARs_Drs(Fbq_ARs(A, Q))
End Function

Private Sub Z_FbqAdoDrs()
Const Fb$ = Samp_Fb_DutyPrepare
Const Q$ = "Select * from Permit"
Drs_XBrw FbqAdoDrs(Fb, Q)
End Sub

Function Fbq_ARs(A$, Q$) As ADODB.Recordset
Set Fbq_ARs = Fb_Cn(A).Execute(Q)
End Function

Sub FbqARun(A$, Q$)
Fb_Cn(A).Execute Q
End Sub

Private Sub Z_FbqARun()
Const Fb$ = Samp_Fb_DutyPrepare
Const Q$ = "Select * into [#a] from Permit"
FbtDrp Fb, "#a"
FbqARun Fb, Q
End Sub


Private Sub Z()
Z_FbqAdoDrs
Z_FbqARun
MAdo_Fb_Fbq:
End Sub
