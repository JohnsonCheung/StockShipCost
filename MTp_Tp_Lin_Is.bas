Attribute VB_Name = "MTp_Tp_Lin_Is"
Option Compare Binary
Option Explicit
Function Lin_IsRmkLin(A) As Boolean

End Function

Function Lin_IsTpRmkLin(A$) As Boolean
Dim L$: L = LTrim(A)
If L <> "" Then
    If XHas_Pfx(L, "--") Then
        Lin_IsTpRmkLin = True
    End If
End If
End Function
