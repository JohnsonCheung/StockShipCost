Attribute VB_Name = "MVb_Str_Tak"
Option Compare Binary
Option Explicit

Function XTak_BefDot$(A)
XTak_BefDot = XTak_Bef(A, ".")
End Function

Function XTak_Aft$(S, Sep, Optional NoTrim As Boolean)
XTak_Aft = Brk1(S, Sep, NoTrim).S2
End Function

Function XTak_AftAt$(A, At&, S)
If At = 0 Then Exit Function
XTak_AftAt = Mid(A, At + Len(S))
End Function

Function XTak_AftDotOrAll$(A)
XTak_AftDotOrAll = XTak_AftOrAll(A, ".")
End Function

Function XTak_AftDot$(A)
XTak_AftDot = XTak_Aft(A, ".")
End Function

Function XTak_AftMust$(A, Sep, Optional NoTrim As Boolean)
XTak_AftMust = Brk(A, Sep, NoTrim).S2
End Function

Function XTak_AftOrAll$(S, Sep, Optional NoTrim As Boolean)
XTak_AftOrAll = Brk2(S, Sep, NoTrim).S2
End Function

Function XTak_AftOrAllRev$(A, S)
XTak_AftOrAllRev = StrDft(XTak_AftRev(A, S), A)
End Function

Function XTak_AftRev$(S, Sep, Optional NoTrim As Boolean)
XTak_AftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function

Function XTak_Bef$(S, Sep, Optional NoTrim As Boolean)
XTak_Bef = Brk2(S, Sep, NoTrim).S1
End Function

Function XTak_BefAt(A, At&)
If At = 0 Then Exit Function
XTak_BefAt = Left(A, At - 1)
End Function

Function XTak_BefDD$(A)
XTak_BefDD = RTrim(XTak_BefOrAll(A, "--"))
End Function

Function XTak_BefDDD$(A)
XTak_BefDDD = RTrim(XTak_BefOrAll(A, "---"))
End Function

Function XTak_BefMust$(S, Sep$, Optional NoTrim As Boolean)
XTak_BefMust = Brk(S, Sep, NoTrim).S1
End Function

Function XTak_BefOrAll$(S, Sep, Optional NoTrim As Boolean)
XTak_BefOrAll = Brk1(S, Sep, NoTrim).S1
End Function

Function XTak_BefOrAllRev$(A, S)
XTak_BefOrAllRev = StrDft(XTak_BefRev(A, S), A)
End Function

Function XTak_BefRev$(A, Sep, Optional NoTrim As Boolean)
XTak_BefRev = Brk2Rev(A, Sep, NoTrim).S1
End Function
Function TakP123(A, S1, S2) As String()
Dim P1&, P2&
P1 = InStr(A, S1)
P2 = InStr(P1 + Len(S1), A, S2)
If P2 > P1 And P1 > 0 And P2 > 0 Then
    PushI TakP123, Left(A, P1)
    Dim L&
        L = P2 - P1 - Len(S1)
    PushI TakP123, Mid(A, P1 + Len(S1), L)
    PushI TakP123, Mid(A, P2 + Len(S2))
End If
End Function
Sub TakP123Asg(A, S1, S2, O1, O2, O3)
AyAsg TakP123(A, S1, S2), O1, O2, O3
End Sub
Private Sub Z_XTak_BefFstLas()
Dim S, Fst, LAs
S = " A_1$ = ""Private Function ZChunk$(ConstLy$(), IChunk%)"" & _"
Fst = vbDblQuote
LAs = vbDblQuote
Ept = "Private Function ZChunk$(ConstLy$(), IChunk%)"
GoSub Tst
Exit Sub
Tst:
    Act = TakBetFstLas(S, Fst, LAs)
    C
    Return
End Sub
Function TakBetFstLas$(S, Fst, LAs)
TakBetFstLas = XTak_BefRev(XTak_Aft(S, Fst), LAs)
End Function

Function TakBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InclMarker Then O = S1 & O & S2
   TakBet = O
End With
End Function

Private Sub Z_XTak_BetBkt()
Dim Act$
   Dim S$
   S = "sdklfjdsf(1234()567)aaa("
   Act = XTak_BetBkt(S)
   Ass Act = "1234()567"
End Sub

Function XTak_Nm$(A)
Dim J%
If Not IsLetter(Left(A, 1)) Then Exit Function
For J = 2 To Len(A)
    If Not IsNmChr(Mid(A, J, 1)) Then
        XTak_Nm = Left(A, J - 1)
        Exit Function
    End If
Next
XTak_Nm = A
End Function

Function XTak_Pfx$(A, Pfx$) ' Return [Pfx] if [Lin] has such pfx else return ""
If XHas_Pfx(Lin, Pfx) Then XTak_Pfx = Pfx
End Function

Function PfxAy_FstSpc$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
Dim P
For Each P In PfxAy
    If XHas_Pfx(Lin, P & " ") Then PfxAy_FstSpc = P: Exit Function
Next
End Function

Function XTak_PfxAy$(A, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
Dim P
For Each P In PfxAy
    If XHas_Pfx(A, P) Then XTak_PfxAy = P: Exit Function
Next
End Function

Function TakSfxAy$(A, SfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
Dim S
For Each S In SfxAy
    If XHas_Sfx(A, S) Then TakSfxAy = S: Exit Function
Next
End Function

Function XTak_PfxAySpc$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
XTak_PfxAySpc = PfxAy_FstSpc$(PfxAy, Lin)
End Function

Function XTak_PfxS$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx+" " else return ""
If XHas_Pfx(Lin, Pfx) Then If Mid(Lin, Len(Pfx) + 1, 1) = " " Then XTak_PfxS = Pfx
End Function

Function TakT1$(A)
If XTak_FstChr(A) <> "[" Then TakT1 = XTak_Bef(A, " "): Exit Function
Dim P%
P = InStr(A, "]")
If P = 0 Then Stop
TakT1 = Mid(A, 2, P - 2)
End Function

Private Sub Z_XTak_AftBkt()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = XTak_AftBkt(A)
    C
    Return
End Sub

Private Sub Z_TakBet()
Dim Lin$
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??       | DATABASE= | ; | ??":            GoSub Tst
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX | DATABASE= | ; | ??":            GoSub Tst
Lin = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf  | Data Source= | ; | Johnson":    GoSub Tst
Exit Sub
Tst:
    Dim FmStr$, ToStr$
    AyAsg AyTrim(SplitVBar(Lin)), Lin, FmStr, ToStr, Ept
    Act = TakBet(Lin, FmStr, ToStr)
    C
    Return
End Sub

Private Sub ZZ_XTak_BetBkt()
Dim A$
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Ept = "AA":        A = "XXX(AA)XX":   GoSub Tst
Ept = "A$()A":     A = "(A$()A)XX":   GoSub Tst
Ept = "O$()":      A = "(O$()) As X": GoSub Tst
Exit Sub
Tst:
    Act = XTak_BetBkt(A)
    C
    Return
End Sub


Private Sub Z()
Z_XTak_AftBkt
Z_XTak_BefFstLas
Z_TakBet
Z_XTak_BetBkt
MVb_Str_Tak:
End Sub
