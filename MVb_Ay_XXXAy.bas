Attribute VB_Name = "MVb_Ay_XXXAy"
Option Compare Binary
Option Explicit
Function AyIntAy(A) As Integer()
AyIntAy = AyInto(A, EmpIntAy)
End Function

Function Ay_XTak_BetBkt(A) As String()
Dim I
For Each I In AyNz(A)
    PushI Ay_XTak_BetBkt, XTak_BetBkt(I)
Next
End Function
Function AyBoolAy(A) As Boolean()
AyBoolAy = AyInto(A, EmpBoolAy)
End Function


Function AyInto(A, OIntoAy)
If TypeName(A) = TypeName(OIntoAy) Then
    AyInto = A
    Exit Function
End If
AyInto = Ay_XCln(OIntoAy)
Dim I
For Each I In AyNz(A)
    Push AyInto, I
Next
End Function

Function AyLngAy(A) As Long()
AyLngAy = AyInto(A, EmpLngAy)
End Function




Function AyBytAy(A) As Byte()
AyBytAy = AyInto(A, EmpBytAy)
End Function


Function AyDblAy(A) As Double()
AyDblAy = AyInto(A, EmpDblAy)
End Function

Function AySngAy(A) As Single()
AySngAy = AyInto(A, EmpSngAy)
End Function


Function AyDteAy(A) As Date()
AyDteAy = AyAsgAy(A, EmpDteAy)
End Function

Private Sub ZZ_AyIntAy()
Dim Act%(): Act = AyIntAy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Function Ay_Sy(A) As String()
Ay_Sy = AyInto(A, EmpSy)
End Function

Function Ay_SyNOBLANK(A) As String()
Dim I
For Each I In AyNz(A)
    PushNonBlankStr Ay_SyNOBLANK, I
Next
End Function


Private Sub Z_AyIntAy()
Dim Act%(): Act = AyIntAy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub


Private Sub Z()
Z_AyIntAy
MVb_Ay_XXXAy:
End Sub
