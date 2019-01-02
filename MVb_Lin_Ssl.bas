Attribute VB_Name = "MVb_Lin_Ssl"
Option Compare Binary
Option Explicit
Function SslAy_Sy(A$()) As String()
Dim O$(), L
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushAy O, Ssl_Sy(L)
Next
SslAy_Sy = O
End Function

Function SslHas(A, N) As Boolean
SslHas = Ay_XHas(Ssl_Sy(A), N)
End Function

Function SslIx&(A, N)
SslIx = Ay_Ix(Ssl_Sy(A), N)
End Function

Function SslJnComma$(Ssl)
SslJnComma = JnComma(Ssl_Sy(Ssl))
End Function

Function SslJnQuoteComma$(Ssl)
SslJnQuoteComma = JnComma(AyQuote(Ssl_Sy(Ssl), "'"))
End Function

Function SslSqBktCsv$(A)
Dim B$(), C$()
B = Ssl_Sy(A)
C = Ay_XQuote_SqBkt(B)
SslSqBktCsv = JnComma(C)
End Function

Function Ssl_Sy(A) As String()
Ssl_Sy = SplitSpc(RplDblSpc(Trim(A)))
End Function
