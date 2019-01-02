Attribute VB_Name = "MDta__ObjPrp"
Option Compare Binary
Option Explicit

Function ItrPrpDrs(A, PrpNy0) As Drs
Dim P$(): P = CvNy(PrpNy0)
Dim Dry()
    Dim I
    For Each I In A
        PushI Dry, Obj_PrpDr(I, P)
    Next
Set ItrPrpDrs = New_Drs(P, Dry)
End Function
Function Oy_Into(A, OInto)
Dim O, J&, Obj
O = Ay_XReSz(OInto, A)
For Each Obj In AyNz(A)
    Set O(J) = Obj
    J = J + 1
Next
End Function
Function Oy_Drs(A, PrpNy0) As Drs
Dim PrpNy$(): PrpNy = CvNy(PrpNy0)
Set Oy_Drs = New_Drs(PrpNy, Oy_Dry(A, PrpNy))
End Function

Function Oy_Dry(A, PrpNy0) As Variant()
Dim O(), U%, I
Dim PrpNy$()
PrpNy = CvNy(PrpNy0)
For Each I In A
    Push O, ObjDr(I, PrpNy)
Next
Oy_Dry = O
End Function

Private Sub Z_ItrPrpDrs()
Drs_XBrw ItrPrpDrs(excel.Application.AddIns, "Name FullName CLSId Installed")
Drs_XBrw ItrPrpDrs(Dbt_Fds(Fb_Db(Samp_Fb_Duty_Dta), "Permit"), "Name Type Required")
'Drs_XBrw ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
End Sub


Private Sub Z()
Z_ItrPrpDrs
MDta__Obj_Prp:
End Sub
