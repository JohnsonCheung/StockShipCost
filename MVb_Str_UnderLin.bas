Attribute VB_Name = "MVb_Str_UnderLin"
Option Compare Binary
Option Explicit
Function UnderLin$(A$)
UnderLin = String(Len(A), "-")
End Function

Function UnderLinDbl$(A)
UnderLinDbl = String(Len(A), "=")
End Function

Function PushMsgUnderLinDbl(O$(), M$)
Push O, M
Push O, UnderLinDbl(M)
End Function

Function PushUnderLin(O$())
Push O, UnderLin(Ay_LasEle(O))
End Function

Function PushUnderLinDbl(O$())
Push O, UnderLinDbl(Ay_LasEle(O))
End Function

Function LinesUnderLin$(A)
LinesUnderLin = StrDup("-", Lines_Wdt(A))
End Function

Function PushMsgUnderLin(O$(), M$)
Push O, M
Push O, UnderLin(M)
End Function
