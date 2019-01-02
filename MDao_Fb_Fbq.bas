Attribute VB_Name = "MDao_Fb_Fbq"
Option Compare Binary
Option Explicit

Private Sub Z_Fbq_Ws()
Ws_XVis Fbq_Ws(Samp_Fb_Duty_Dta, "Select * from KE24")
End Sub

Function Fbq_Ws(A, Sql) As Worksheet
Set Fbq_Ws = Drs_Ws(Fbq_Drs(A, Sql))
End Function

Function Fbq_Drs(A, Sql) As Drs
Set Fbq_Drs = ARs_Drs(Fb_Cn(A).Execute(Sql))
End Function

Sub FbqRun(A, Sql)
Fb_Cn(A).Execute Sql
End Sub
