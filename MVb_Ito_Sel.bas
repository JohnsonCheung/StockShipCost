Attribute VB_Name = "MVb_Ito_Sel"
Option Compare Binary
Option Explicit

Function ItoSel(A, PrpNy0) As Variant()
Dim Dry(), O, Dr(), N, PrpNy$()
PrpNy = CvNy(PrpNy0)
For Each O In A
    Erase Dr
    For Each N In PrpNy
        PushI Dr, Obj_Prp(O, N)
    Next
    PushI Dry, Dr
Next
ItoSel = Dry
End Function

Function ItoDry(A, PrpNy0) As Variant()
ItoDry = ItoSel(A, PrpNy0)
End Function
