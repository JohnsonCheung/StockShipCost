Attribute VB_Name = "MDta___Fun"
Option Compare Binary
Option Explicit

Function Ap_DtAy(ParamArray Ap()) As Dt()
Dim Av(): Av = Ap
Ap_DtAy = AyInto(Av, EmpDtAy)
End Function


Function LinSimTyAyDr(A, B() As eSimTy) As Variant()
End Function
Private Sub ZZ_ItrPrpDrs()
Drs_XDmp ItrPrpDrs(Application.Vbe.VBProjects, "Name Type")
End Sub

Function SqAlign(Sq(), W%()) As Variant()
If UBound(Sq, 2) <> Sz(W) Then Stop
Dim C%, R%, Wdt%, O
O = Sq
For C = 1 To UBound(Sq, 2) - 1 ' The last column no need to align
    Wdt = W(C - 1)
    For R = 1 To UBound(Sq, 1)
        O(R, C) = XAlignL(Sq(R, C), Wdt)
    Next
Next
SqAlign = O
End Function

