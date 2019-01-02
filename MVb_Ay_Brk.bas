Attribute VB_Name = "MVb_Ay_Brk"
Option Compare Binary
Option Explicit
Function Ay_XBrk_BY_PFX(A, Pfx, Optional CmpMth As VbCompareMethod = VbCompareMethod.vbTextCompare) As AyPair
Dim O As AyPair
O.A = Ay_XCln(A)
O.B = O.A
Dim V
For Each V In AyNz(A)
    If XHas_Pfx(V, Pfx, CmpMth) Then
        PushI O.B, V
    Else
        PushI O.A, V
    End If
Next
Ay_XBrk_BY_PFX = O
End Function

Function Ay_XBrk_BY_ELE(A, Ele) As AyPair
Dim O As AyPair
O.A = Ay_XCln(A)
O.B = O.A
Dim J%
For J = 0 To UB(A)
    If A(J) = Ele Then Exit For
    PushI O.A, A(J)
Next
For J = J + 1 To UB(A)
    PushI O.B, A(J)
Next
Ay_XBrk_BY_ELE = O
End Function

Function Ay_XBrk_ABC(A, FmIx&, ToIx&) As AyABC
Dim O As AyABC
O.A = Ay_XWh_FmTo(A, 0, FmIx - 1)
O.B = Ay_XWh_FmTo(A, FmIx, ToIx)
O.C = Ay_XWh_Fm(A, ToIx + 1)
Ay_XBrk_ABC = O
End Function

Function Ay_XBrk_ABC_FTIx(A, B As FTIx) As AyABC
Ay_XBrk_ABC_FTIx = Ay_XBrk_ABC(A, B.FmIx, B.ToIx)
End Function
