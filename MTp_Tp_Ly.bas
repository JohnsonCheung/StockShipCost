Attribute VB_Name = "MTp_Tp_Ly"
Option Compare Binary
Option Explicit
Function LyHasMajPfx(A$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(A)
    If XHas_Pfx(A(J), MajPfx) Then Cnt = Cnt + 1
Next
LyHasMajPfx = Cnt > (Sz(A) \ 2)
End Function
