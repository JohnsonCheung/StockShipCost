Attribute VB_Name = "MTp_X_RRCC"
Option Compare Binary
Option Explicit

Function RRCC_IsEmp(A As RRCC) As Boolean
RRCC_IsEmp = True
With A
   If .R1 <= 0 Then Exit Function
   If .R2 <= 0 Then Exit Function
   If .R1 > .R2 Then Exit Function
End With
RRCC_IsEmp = False
End Function

Function CvRRCC(A) As RRCC
Set CvRRCC = A
End Function
Function New_RRCC(R1, R2, C1, C2) As RRCC
Set New_RRCC = New RRCC
With New_RRCC
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function
