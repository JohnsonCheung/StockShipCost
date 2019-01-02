Attribute VB_Name = "MVb_Str_SubStr"
Option Compare Binary
Option Explicit

Function XTak_LasChr$(A)
XTak_LasChr = Right(A, 1)
End Function

Function XTak_FstChr$(A)
XTak_FstChr = Left(A, 1)
End Function

Function XTak_FstTwoChr$(A)
XTak_FstTwoChr = Left(A, 2)
End Function
