Attribute VB_Name = "MVb_Dic_Wh"
Option Compare Binary
Option Explicit
Function DicWh(A As Dictionary, Ky0) As Dictionary
Set DicWh = New Dictionary
Dim K
For Each K In CvNy(Ky0)
    If A.Exists(K) Then
        DicWh.Add K, A(K)
    End If
Next
End Function
