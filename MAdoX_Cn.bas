Attribute VB_Name = "MAdoX_Cn"
Option Compare Binary
Option Explicit
Function Fx_Cn(A) As ADODB.Connection
Set Fx_Cn = CnStr_Cn(Fx_CnStr_Ado(A))
End Function

Function Fb_Cn(A) As ADODB.Connection
Set Fb_Cn = CnStr_Cn(Fb_CnStr_Ado(A))
End Function

Private Sub Z_Fb_Cn()
Dim Cn
Set Cn = Fb_Cn(Samp_Fb_Duty_Dta)
Stop
End Sub


