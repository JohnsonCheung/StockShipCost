Attribute VB_Name = "MVb_Ay_ReOrder"
Option Compare Binary
Option Explicit
Function AyReOrd(A, PartialIxAy)
Dim Ay, Ix
    Ay = Ay_XCln(A)
    For Each Ix In PartialIxAy
        PushI Ay, A(Ix)
    Next
AyReOrd = AyReOrdAy(A, Ay)
End Function

Function AyReOrdAy(A, SubAy)
If Not Ay_XHasAy(A, SubAy) Then Stop
AyReOrdAy = Ay_XAdd_(SubAy, AyMinus(A, SubAy))
End Function
