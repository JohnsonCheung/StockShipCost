Attribute VB_Name = "MAdo_Fds"
Option Compare Binary
Option Explicit
Function AFds_Dr(A As ADODB.Fields) As Variant()
Dim F As ADODB.Field
For Each F In A
   PushI AFds_Dr, F.Value
Next
End Function

Function AFds_Fny(A As ADODB.Fields) As String()
AFds_Fny = Itr_Ny(A)
End Function
