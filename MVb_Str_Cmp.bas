Attribute VB_Name = "MVb_Str_Cmp"
Option Compare Binary
Option Explicit

Sub Lines_XCmp(A, B, Optional N1$ = "A", Optional N2$ = "B")
Brw Lines_CmpFmt(A, B, N1, N2)
End Sub

Function Lines_CmpFmt(A, B, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$) As String()
If A = B Then Exit Function
Dim AA$(), BB$()
AA = SplitCrLf(A)
BB = SplitCrLf(B)
If Ay_IsEq(AA, BB) Then Exit Function
Dim DifAt&
    DifAt = AyAB_DIfIx(AA, BB)
Dim O$(), J&, MinU&
    PushNonBlankStr O, Hdr
    PushI O, QQ_Fmt("LinesCnt=? (?)", Sz(AA), N1)
    PushI O, QQ_Fmt("LinesCnt=? (?)", Sz(BB), N2)
    For J = 0 To DifAt - 1
        PushI O, J & ":" & AA(J)
    Next
    MinU = Min(UB(AA), UB(BB))
    PushIAy O, Lines_CmpFmt1(AA, BB, MinU, DifAt)
    PushIAy O, Lines_CmpFmt2(AA, BB, MinU, DifAt)
    PushI O, N1 & " " & String(10, "-")
    PushIAy O, AA
    PushI O, N2 & " " & String(10, "-")
    PushIAy O, BB
Lines_CmpFmt = O
End Function

Private Function Lines_CmpFmt1(A$(), B$(), MinU&, DifAt&) As String()
Dim Pfx$, J&
For J = DifAt To MinU
    Pfx = J & ":"
    PushI Lines_CmpFmt1, Pfx & A(J) & "<Len(" & Len(A(J)) & ")"
    PushI Lines_CmpFmt1, Space(Len(Pfx)) & B(J) & "<Len(" & Len(B(J)) & ")"
Next
End Function

Private Function Lines_CmpFmt2(A$(), B$(), MinU&, DifAt&) As String()
Dim X$(): X = IIf(UB(A) > UB(B), A, B)
Dim J&
For J = MinU + 1 To UB(X)
    PushI Lines_CmpFmt2, J & ":" & X(J)
Next
End Function

Sub Str_XCmp(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$)
If A = B Then Exit Sub
Brw Str_CmpFmt(A, B, N1, N2, Hdr)
End Sub

Function Str_CmpFmt(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$) As String()
If IsLines(A) Or IsLines(B) Then Str_CmpFmt = Lines_CmpFmt(A, B, N1, N2, Hdr): Exit Function
If A = B Then Exit Function
Dim DifAt&
    DifAt = Str_DifPos(A, B)
Dim O$()
    PushI O, QQ_Fmt("Str-(?)-Len: ?", N1, Len(A))
    PushI O, QQ_Fmt("Str-(?)-Len: ?", N2, Len(B))
    PushI O, "Dif At: " & DifAt
    PushIAy O, Len_LblAy(Max(Len(A), Len(B)))
    PushI O, A
    PushI O, B
    PushI O, Space(DifAt - 1) & "^"
Str_CmpFmt = O
End Function

Function Str_DifPos&(A, B)
Dim O&
For O = 1 To Min(Len(A), Len(B))
    If Mid(A, O, 1) <> Mid(B, O, 1) Then Str_DifPos = O: Exit Function
Next
End Function

Function AyAB_DIfIx&(A, B)
Dim O&
For O = 0 To Min(Sz(A), Sz(B))
    If A(O) <> B(O) Then AyAB_DIfIx = O: Exit Function
Next
End Function

Function Len_LblAy(L&) As String()
If L <= 0 Then XThw CSub, "Length should be >0", "Length", L
Dim N%
    N = NDig(L)
PushNonBlankStr Len_LblAy, Len_LblLin1(L)
PushI Len_LblAy, Len_LblLin2(L)
End Function

Private Function Len_LblLin1$(L&)
Dim J&, O$(), N&
PushI O, Space(9)
For J = 1 To (L - 1) \ 10 + 1
    N = J * 10
    PushI O, N & Space(10 - NDig(N))
Next
Len_LblLin1 = Join(O, "")
End Function

Private Function Len_LblLin2$(L&)
Dim Q&, R%
Const C$ = "123456789 "
Q = (L - 1) \ 10 + 1
R = (L - 1) Mod 10 + 1
Len_LblLin2 = StrDup(C, Q) & Left(C, R)
End Function

Private Sub Z_Lines_CmpFmt()
Dim A$, B$
A = XRpl_VBar("AAAAAAA|bbbbbbbb")
B = XRpl_VBar("AAAAAAA|bbbbbbbb ")
GoSub Tst
Exit Sub
Tst:
    Act = Lines_CmpFmt(A, B)
    Brw Act
    Return

End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$
End Sub

Private Sub Z()
End Sub
