Attribute VB_Name = "MDta_Sel"
Option Compare Binary
Option Explicit
Function DrSel(A, IxAy) As Variant()
Dim Ix
For Each Ix In IxAy
    PushI DrSel, A(Ix)
Next
End Function
Function DrySel(A, IxAy) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    PushI DrySel, DrSel(Dr, IxAy)
Next
End Function
Function DrySelIxAp(A, ParamArray IxAp()) As Variant()
Dim IxAy(): IxAy = IxAp
DrySelIxAp = DrySel(A, IxAy)
End Function

Function Drs_XSel(A As Drs, FF) As Drs
Dim Fny$(): Fny = CvNy(FF)
If Ay_IsEq(A.Fny, Fny) Then Set Drs_XSel = A: Exit Function
Ay_XBrw Ay_XHasAyChk(A.Fny, Fny)

Dim Ix&(): Ix = Ay_IxAy(A.Fny, Fny)
Dim Dry(), Dr
For Each Dr In A.Dry
    PushI Dry, DrSel(Dr, Ix)
Next
Set Drs_XSel = New_Drs(Fny, Dry)
End Function

Private Sub Z_Drs_XSel()
'Drs_XBrw Drs_XSel(Vmd.MthDrs, "MthNm Mdy Ty MdNm")
'Drs_XBrw Vmd.MthDrs
End Sub

Function DtSel(A As Dt, ColNy0) As Dt
Dim ReOrdFny$(): ReOrdFny = CvNy(ColNy0)
Dim IxAy&(): IxAy = Ay_IxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
Set DtSel = New_Dt(A.DtNm, OFny, ODry)
End Function


Private Sub Z()
Z_Drs_XSel
MDta_Sel:
End Sub
