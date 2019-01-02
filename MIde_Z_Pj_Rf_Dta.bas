Attribute VB_Name = "MIde_Z_Pj_Rf_Dta"
Option Compare Binary
Option Explicit
Private Sub Z_PjRfDrs()
Drs_XDmp PjRfDrs(CurPj)
End Sub
Sub PjRfDrs_XDmp(A As VBProject)
Drs_XDmp PjRfDrs(A)
End Sub
Function PjRfDrs(A As VBProject) As Drs
Set PjRfDrs = New_Drs(PjRfFny, PjRfDry(A))
End Function
Property Get PjRfFny() As String()
PjRfFny = ItmAddAy("Pj", RfFny)
End Property
Function PjRfDry(A As VBProject) As Variant()
Dim R As VBIDE.Reference, N$
N = A.Name
For Each R In A.References
    PushI PjRfDry, ItmAddAy(N, RfDr(R))
Next
End Function
Function RfDr(A As VBIDE.Reference) As Variant()
With A
RfDr = Array(.Name, .GUID, .Major, .Minor, .FullPath, .Description, .BuiltIn, .Type, .IsBroken)
End With
End Function
Property Get RfFny() As String()
RfFny = Ssl_Sy(RmvDotComma(".Name, .GUID, .Major, .Minor, .FullPath, .Description, .BuiltIn, .Type, .IsBroken"))
End Property

Function PjAyRfDrs(A() As VBProject) As Drs
Dim P
For Each P In AyNz(A)
    Drs_Push PjAyRfDrs, PjRfDrs(CvPj(P))
Next
End Function
