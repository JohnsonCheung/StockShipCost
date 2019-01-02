Attribute VB_Name = "MVb_Str_Macro"
Option Compare Binary
Option Explicit
Function MacroNy(A, Optional ExlBkt As Boolean, Optional OpnBkt$ = vbOpnSqBkt) As String()
'MacroStr-A is a with ..[xx].., this sub is to return all xx
Dim Q1$, Q2$
    Q1 = OpnBkt
    Q2 = OpnBkt_ClsBktt(OpnBkt)
If Not HasSubStr(A, Q1) Then Exit Function

Dim Ay$(): Ay = Split(A, Q1)
Dim O$(), J%
For J = 1 To UB(Ay)
    Push O, XTak_Bef(Ay(J), Q2)
Next
If Not ExlBkt Then
    O = Ay_XAdd_PfxSfx(O, Q1, Q2)
End If
MacroNy = O
End Function
