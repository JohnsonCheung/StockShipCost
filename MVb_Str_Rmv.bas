Attribute VB_Name = "MVb_Str_Rmv"
Option Compare Binary
Option Explicit

Function RmvDotComma$(A)
RmvDotComma = Replace(Replace(A, ",", ""), ".", "")
End Function
Function Rmv2Dash$(A)
Rmv2Dash = RTrim(RmvAft(A, "--"))
End Function

Function Rmv3Dash$(A)
Rmv3Dash = RTrim(RmvAft(A, "---"))
End Function

Function Rmv3T$(A$)
Rmv3T = RmvTT(RmvT1(A))
End Function

Function RmvAft$(A, Sep$)
RmvAft = Brk1(A, Sep, NoTrim:=True).S1
End Function

Function RmvDDRmk$(A$)
Dim S$
If LinHasDDRmk(A) Then
    S = ""
Else
    S = A
End If
End Function

Function RmvDblSpc$(A)
Dim O$: O = A
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function XRmv_FstChr$(A)
XRmv_FstChr = Mid(A, 2)
End Function

Function RmvFstLasChr$(A)
RmvFstLasChr = XRmv_FstChr(XRmv_LasChr(A))
End Function

Function RmvFstNChr$(A, Optional N% = 1)
RmvFstNChr = Mid(A, N + 1)
End Function

Function XRmv_FstNonLetter$(A)
If Asc_IsLetter(Asc(A)) Then
    XRmv_FstNonLetter = A
Else
    XRmv_FstNonLetter = XRmv_FstChr(A)
End If
End Function

Function XRmv_LasChr$(A)
XRmv_LasChr = XRmv_LasNChr(A, 1)
End Function

Function XRmv_LasNChr$(A, N%)
XRmv_LasNChr = Left(A, Len(A) - N)
End Function

Function XRmv_Nm$(A)
Dim O%
If Not Asc_IsFstNmChr(Asc(XTak_FstChr(A))) Then GoTo X
For O = 1 To Len(A)
    If Not Asc_IsNmChr(Asc(Mid(A, O, 1))) Then GoTo X
Next
X:
    If O > 0 Then XRmv_Nm = Mid(A, O): Exit Function
    XRmv_Nm = A
End Function

Function XRmv_OptSqBkt$(A)
If Not HasSqBkt(A) Then XRmv_OptSqBkt = A: Exit Function
XRmv_OptSqBkt = RmvFstLasChr(A)
End Function

Function XRmv_Pfx$(A, Pfx)
If XHas_Pfx(A, Pfx) Then XRmv_Pfx = Mid(A, Len(Pfx) + 1) Else XRmv_Pfx = A
End Function

Function XRmv_PfxAy$(A, PfxAy)
Dim Pfx
For Each Pfx In PfxAy
    If XHas_Pfx(A, CStr(Pfx)) Then XRmv_PfxAy = XRmv_Pfx(A, Pfx): Exit Function
Next
XRmv_PfxAy = A
End Function
Function XRmv_PfxSpc$(A, Pfx)
If Not XHas_PfxSpc(A, Pfx) Then XRmv_PfxSpc = A: Exit Function
XRmv_PfxSpc = LTrim(Mid(A, Len(Pfx) + 2))
End Function
Function XRmv_PfxAySpc$(A, PfxAy)
Dim P
For Each P In PfxAy
    If XHas_PfxSpc(A, P) Then
        XRmv_PfxAySpc = LTrim(Mid(A, Len(P) + 2))
        Exit Function
    End If
Next
XRmv_PfxAySpc = A
End Function

Function RmvSfx$(A, Sfx)
If XHas_Sfx(A, Sfx) Then RmvSfx = Left(A, Len(A) - Len(Sfx)) Else RmvSfx = A
End Function

Function XRmv_SngQuote$(A)
If Not IsSngQuoted(A) Then XRmv_SngQuote = A: Exit Function
XRmv_SngQuote = RmvFstLasChr(A)
End Function

Function XRmv_SqBkt$(A)
If Not HasSqBkt(A) Then XRmv_SqBkt = A: Exit Function
XRmv_SqBkt = RmvFstLasChr(A)
End Function

Function RmvT1$(A)
RmvT1 = Lin_1TRst(A)(1)
End Function

Function RmvTT$(A)
RmvTT = RmvT1(RmvT1(A))
End Function

Function RmvUSfx$(A) ' Return upcase Sfx
Dim J%, Fnd As Boolean, C%
For J = Len(A) To 2 Step -1 ' don't find the first char if non-UCase, to use 'To 2'
    C = Asc(Mid(A, J, 1))
    If Not Asc_IsUCase(C) Then
        Fnd = True
        Exit For
    End If
Next
If Fnd Then
    RmvUSfx = Left(A, J)
Else
    RmvUSfx = A
End If
End Function

Private Sub Z_RmvT1()
Ass RmvT1("  df dfdf  ") = "dfdf"
End Sub


Private Sub Z_XRmv_Nm()
Dim Nm$
Nm = "lksdjfsd f"
Ept = " f"
GoSub Tst
Exit Sub
Tst:
    Act = XRmv_Nm(Nm)
    C
    Return
End Sub

Private Sub Z_XRmv_Pfx()
Ass XRmv_Pfx("aaBB", "aa") = "BB"
End Sub

Private Sub Z_XRmv_PfxAy()
Dim A$, PfxAy$()
PfxAy = Ssl_Sy("ZZ_ Z_"): Ept = "ABC"
A = "Z_ABC": GoSub Tst
A = "ZZ_ABC": GoSub Tst
Exit Sub
Tst:
    Act = XRmv_PfxAy(A, PfxAy)
    C
    Return
End Sub


Private Sub Z()
Z_XRmv_Nm
Z_XRmv_Pfx
Z_XRmv_PfxAy
Z_RmvT1
MVb_Str_Rmv:
End Sub
