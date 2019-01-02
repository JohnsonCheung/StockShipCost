Attribute VB_Name = "MVb_Csv"
Option Compare Binary
Option Explicit

Function CvCsv$(A)
Select Case True
Case IsStr(A): CvCsv = """" & A & """"
Case IsDte(A): CvCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
Case TypeName(A) = "RecordSet2": CvCsv = "*Att"
Case Else: CvCsv = IIf(IsNull(A), "", A)
End Select
End Function

Function AyCsv$(A)
If Sz(A) = 0 Then Exit Function
Dim O$(), U&, J&, V
U = UB(A)
ReDim O(U)
For Each V In A
    O(J) = CvCsv(V)
    J = J + 1
Next
AyCsv = Join(O, ",")
End Function
