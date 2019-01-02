Attribute VB_Name = "MDao_Schm_Std"
Option Compare Binary
Option Explicit

Function EleNm_IsStd(A) As Boolean
Stop '
End Function

Function FldNm_IsStd(A) As Boolean
FldNm_IsStd = True
If A = "CrtDte" Then Exit Function
If Ay_XHas(Ssl_Sy("Id Ty Nm"), Right(A, 2)) Then Exit Function
If Ay_XHas(Ssl_Sy("Dte Amt"), Right(A, 3)) Then Exit Function
FldNm_IsStd = False
End Function

