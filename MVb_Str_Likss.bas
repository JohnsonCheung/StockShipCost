Attribute VB_Name = "MVb_Str_Likss"
Option Compare Binary
Option Explicit
Function XLik_Likss(S, Likss) As Boolean
XLik_Likss = XLik_LikAy(S, Ssl_Sy(Likss))
End Function

Function XLik_LikAy(S, LikAy$()) As Boolean
Dim I
For Each I In AyNz(LikAy)
    If S Like I Then XLik_LikAy = True: Exit Function
Next
End Function

Function XLik_LikssAy(S, LikssAy) As Boolean
Dim Likss
For Each Likss In LikssAy
    If XLik_Likss(S, Likss) Then XLik_LikssAy = True: Exit Function
Next
End Function

