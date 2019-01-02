Attribute VB_Name = "MIde__CdLin"
Option Compare Binary
Option Explicit

Function Ay_XWh_CdLin(A) As String()
Dim L
For Each L In AyNz(A)
    If IsCdLin(L) Then
        PushI Ay_XWh_CdLin, L
    End If
Next
End Function

Function IsCdLin(A) As Boolean
Dim L$: L = Trim(A)
If A = "" Then Exit Function
If Left(A, 1) = "'" Then Exit Function
IsCdLin = True
End Function

Private Sub ZZ()
Dim A As Variant
Ay_XWh_CdLin A
IsCdLin A
End Sub

Private Sub Z()
End Sub
