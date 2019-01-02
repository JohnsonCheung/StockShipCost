Attribute VB_Name = "MVb_Is_Asc"
Option Compare Binary
Option Explicit

Function Asc_IsDig(A%) As Boolean
Asc_IsDig = &H30 <= A And A <= &H39
End Function

Function Asc_IsDigit(A%) As Boolean
If A < 48 Then Exit Function
If A > 57 Then Exit Function
Asc_IsDigit = True
End Function

Function Asc_IsFstNmChr(A%) As Boolean
Asc_IsFstNmChr = Asc_IsLetter(A)
End Function

Function Asc_IsLCase(A%) As Boolean
If A < 97 Then Exit Function
If A > 122 Then Exit Function
Asc_IsLCase = True
End Function

Function Asc_IsLetter(A%) As Boolean
Asc_IsLetter = True
If Asc_IsUCase(A) Then Exit Function
If Asc_IsLCase(A) Then Exit Function
Asc_IsLetter = False
End Function

Function Asc_IsNmChr(A%) As Boolean
Asc_IsNmChr = True
If Asc_IsLetter(A) Then Exit Function
If Asc_IsDig(A) Then Exit Function
Asc_IsNmChr = A = 95 '_
End Function

Function Asc_IsUCase(A%) As Boolean
If A < 65 Then Exit Function
If A > 90 Then Exit Function
Asc_IsUCase = True
End Function
