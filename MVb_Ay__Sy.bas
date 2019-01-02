Attribute VB_Name = "MVb_Ay__Sy"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Ay__Sy."

Function CvNy(Ny0) As String()
Const CSub$ = CMod & "CvNy"
Select Case True
Case IsMissing(Ny0) Or IsEmpty(Ny0)
Case IsStr(Ny0): CvNy = Lin_TermAy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = Ay_Sy(Ny0)
Case Else: XThw CSub, "Given Ny0 must be Missing | Empty | Str | Sy | Ay", "TypeName-Ny0", TypeName(Ny0)
End Select
End Function

Function CvSy(A) As String()
Select Case True
Case IsEmpty(A) Or IsMissing(A)
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = Ay_Sy(A)
Case Else: CvSy = Ap_Sy(CStr(A))
End Select
End Function


Private Sub ZZ()
Dim A
Dim B()
Dim C$
Dim D$()
Dim XX
CvNy A
CvSy A
Ap_Sy B
End Sub

Private Sub Z()
End Sub
