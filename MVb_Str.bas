Attribute VB_Name = "MVb_Str"
Option Compare Binary
Option Explicit
Function Val_XAdd_Lbl$(A, Lbl$)
Dim B$
If IsDate(A) Then
    B = Dte_DTim(CDate(A))
Else
    B = Replace(Replace(A, ";", "%3B"), "=", "%3D")
End If
If A <> "" Then Val_XAdd_Lbl = Lbl & "=" & B
End Function

Function XPad0$(N, NDig%)
XPad0 = Format(N, StrDup("0", NDig))
End Function

Function Str_XApp$(A, L)
If A = "" Then Str_XApp = L: Exit Function
Str_XApp = A & " " & L
End Function

Function XBrk1(A, Sep$, Optional NoTrim As Boolean) As String()
XBrk1 = XBrk1At(A, InStr(A, Sep), Sep, NoTrim)
End Function

Function XBrk1At(A, At&, Sep, Optional NoTrim As Boolean) As String()
Dim O1$, O2$
If At = 0 Then
    O1 = A
Else
    O1 = Left(A, At - 1)
    O2 = Mid(A, At + Len(Sep))
End If
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
XBrk1At = Ap_Sy(O1, O2)
End Function

Function XBrk(A, Sep$, Optional NoTrim As Boolean) As String()
XBrk = XBrkAt(A, InStr(A, Sep), Sep, NoTrim)
End Function

Function XBrkAt(A, At&, Sep, Optional NoTrim As Boolean) As String()
If At = 0 Then Stop
Dim O1$, O2$
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
XBrkAt = Ap_Sy(O1, O2)
End Function

Sub StrBrw(A, Optional Fnn$)
Dim T$: T = TmpFt("StrBrw", Fnn$)
Str_XWrt A, T
Ft_XBrw T
End Sub

Function StrDft$(A, B)
StrDft = IIf(A = "", B, A)
End Function

Function StrDup$(S, N)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrInSfxAy(A, SfxAy$()) As Boolean
StrInSfxAy = Ay_XHasPredPXTrue(SfxAy, "XHas_Sfx", A)
End Function

Function StrMatchPfxAy(A, PfxAy$()) As Boolean
If Sz(PfxAy) = 0 Then Exit Function
Dim Pfx
For Each Pfx In PfxAy
    If A Like Pfx & "*" Then StrMatchPfxAy = True: Exit Function
Next
End Function

Sub Str_XWrt(A, Ft, Optional OvrWrt As Boolean)
If OvrWrt Then Ffn_XDltIfExist (Ft)
Fso.CreateTextFile(Ft, True).Write A
End Sub

Function SubStrCnt&(A, SubStr)
Dim P&: P = 1
Dim L%: L = Len(SubStr)
Dim O%
While P > 0
    P = InStr(P, A, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Function
    O = O + 1
    P = P + L
Wend
SubStrCnt = O
End Function

Function SubStrPos(A, SubStr$) As FTIx
Dim FmIx&: FmIx = InStr(A, SubStr)
Dim ToIx&
If FmIx > 0 Then ToIx = FmIx + Len(SubStr)
SubStrPos = New_FTIx(FmIx, ToIx)
End Function

Private Sub Z_SubStrCnt()
Dim A$, SubStr$
A = "aaaa":                 SubStr = "aa":  Ept = CLng(2): GoSub Tst
A = "aaaa":                 SubStr = "a":   Ept = CLng(4): GoSub Tst
A = "skfdj skldfskldf df ": SubStr = " ":   Ept = CLng(3): GoSub Tst
Exit Sub
Tst:
    Act = SubStrCnt(A, SubStr)
    C
    Return
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C%
Dim D As Boolean
Dim E&
Dim F$()
Val_XAdd_Lbl A, B
XPad0 A, C
Str_XApp A, A
XBrk A, B, D
XBrk1 A, B, D
XBrk1At A, E, A, D
XBrkAt A, E, A, D
StrBrw A, B
StrDft A, A
StrDup A, A
StrInSfxAy A, F
StrMatchPfxAy A, F
Str_XWrt A, A, D
SubStrCnt A, A
SubStrPos A, B
End Sub

Private Sub Z()
Z_SubStrCnt
End Sub
