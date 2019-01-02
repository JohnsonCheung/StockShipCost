Attribute VB_Name = "MVb_Str_Ens"
Option Compare Binary
Option Explicit
Function XEns_Sfx$(A, Sfx$)
If XHas_Sfx(A, Sfx) Then XEns_Sfx = A: Exit Function
XEns_Sfx = A & Sfx
End Function

Function XEns_SfxDot$(A)
XEns_SfxDot = XEns_Sfx(A, ".")
End Function

Function XEns_SfxSC$(A)
XEns_SfxSC = XEns_Sfx(A, ";")
End Function
