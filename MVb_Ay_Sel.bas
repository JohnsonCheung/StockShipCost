Attribute VB_Name = "MVb_Ay_Sel"
Option Compare Binary
Option Explicit

Function AySelT1(A) As String()
Dim I
For Each I In AyNz(A)
    PushI AySelT1, Lin_T1(A)
Next
End Function
