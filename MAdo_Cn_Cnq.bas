Attribute VB_Name = "MAdo_Cn_Cnq"
Option Compare Binary
Option Explicit

Sub CnSqy_XRun(A As ADODB.Connection, Sqy$())
Dim Q
For Each Q In AyNz(Sqy)
   A.Execute Q
Next
End Sub

Private Sub Z_Cnq_Drs()
Dim Cn As ADODB.Connection: Set Cn = Fx_Cn(Samp_Fx_KE24)
Dim Q$: Q = "Select * from [Sheet1$]"
Drs_Ws Cnq_Drs(Cn, Q)
End Sub

Function Cnq_ARs(A As ADODB.Connection, Q) As ADODB.Recordset
Set Cnq_ARs = A.Execute(Q)
End Function

Function Cnq_Drs(A As ADODB.Connection, Q) As Drs
Set Cnq_Drs = ARs_Drs(Cnq_ARs(A, Q))
End Function


Private Sub Z()
Z_Cnq_Drs
MAdo_Cn_Cnq:
End Sub
