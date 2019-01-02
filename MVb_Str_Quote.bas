Attribute VB_Name = "MVb_Str_Quote"
Option Compare Binary
Option Explicit
Function XQuote_Bkt$(A)
XQuote_Bkt = "(" & A & ")"
End Function
Function XQuote$(A, QuoteStr$)
With BrkQuote(QuoteStr)
    XQuote = .S1 & A & .S2
End With
End Function

Function XQuote_Dbl$(A$)
XQuote_Dbl = """" & Replace(A, """", """""") & """"
End Function

Function XQuote_Sng$(A)
XQuote_Sng = "'" & Replace(A, "'", "''") & "'"
End Function

Function XQuote_Dte$(A)
XQuote_Dte = "#" & A & "#"
End Function

Function XQuote_SqBkt$(A)
XQuote_SqBkt = "[" & A & "]"
End Function

Function XQuote_SqBkt_IfNeed$(A)
If IsSqBktNeed(A) Then
    XQuote_SqBkt_IfNeed = "[" & A & "]"
Else
    XQuote_SqBkt_IfNeed = A
End If
End Function
