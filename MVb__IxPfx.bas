Attribute VB_Name = "MVb__IxPfx"
Option Compare Binary
Option Explicit

Function Ay_XAdd_IxPfx(A, Optional BegFm&) As String()
Dim I, J&, N%
J = BegFm
N = Len(CStr(Sz(A)))
For Each I In AyNz(A)
    PushI Ay_XAdd_IxPfx, XAlignR(J, N) & ": " & I
    J = J + 1
Next
End Function

Function LinesAddIxPfx$(A, Optional BegFm&)
LinesAddIxPfx = JnCrLf(Ay_XAdd_IxPfx(SplitCrLf(A), BegFm))
End Function
