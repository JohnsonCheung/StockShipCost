Attribute VB_Name = "MVb_X_S1S2_Fmt"
Option Compare Binary
Option Explicit

Function S1S2Ay_Fmt(A() As S1S2, Optional Nm1$, Optional Nm2$) As String()
If Any_Lines(A) Then
    Dim W1%: W1 = ZW1(A, Nm1)
    Dim W2%: W2 = ZW2(A, Nm2)
    Dim H$: H = WdtAyHdrLin(Ap_IntAy(W1, W2))
    S1S2Ay_Fmt = WFmt(A, H, W1, W2, Nm1, Nm2)
    Exit Function
End If
S1S2Ay_Fmt = WFmtX(A, Nm1, Nm2)
End Function

Function WdtAyHdrLin$(A%())
Dim O$(), W
For Each W In A
    Push O, StrDup("-", W + 2)
Next
WdtAyHdrLin = "|" + Join(O, "|") + "|"
End Function

Private Sub X(O1$(), O2$())
Erase O1
Erase O2
Dim A1$, A2$
A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df":          GoSub X
A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub X
A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df":               GoSub X
A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df":           GoSub X
A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df":            GoSub X
Exit Sub
X:
    PushI O1, XRpl_VBar(A1)
    PushI O2, XRpl_VBar(A2)
    Return
End Sub

Private Function WFmt(A() As S1S2, H$, W1%, W2%, Nm1$, Nm2$) As String()
PushIAy WFmt, WFmt1(H, Nm1, Nm2, W1, W2)
Dim I&
PushI WFmt, H
For I = 0 To UB(A)
   PushIAy WFmt, WFmt2(A(I), W1, W2)
   PushI WFmt, H
Next
End Function

Private Function WFmt1(H$, Nm1$, Nm2$, W1%, W2%) As String()
If Nm1 = "" And Nm2 = "" Then Exit Function
PushI WFmt1, H
PushI WFmt1, "| " & XAlignL(Nm1, W1) & " | " & XAlignL(Nm2, W2) & " |"
End Function

Private Function WFmt2(A As S1S2, W1%, W2%) As String()
Dim S1$(), S2$(), U%
S1 = SplitCrLf(A.S1)
S2 = SplitCrLf(A.S2)
U = Max(UB(S1), UB(S2))
S1 = WFmt3(S1, U, W1)
S2 = WFmt3(S2, U, W2)
Dim J&
For J = 0 To U
    PushI WFmt2, "| " & S1(J) & " | " & S2(J) & " |"
Next
End Function

Private Function WFmt3(Ay$(), U%, W%) As String()
ReDim Preserve Ay(U)
Dim I
For Each I In AyNz(Ay)
    PushI WFmt3, XAlignL(I, W)
Next
End Function

Private Function WFmtX(A() As S1S2, Nm1$, Nm2$) As String()
If Sz(A) = 0 Then Exit Function
Dim S1$(), Sep$, S2$()
    S1 = Ay_XAlign_L(S1S2Ay_Sy1(A))
    S2 = S1S2Ay_Sy2(A)
Sep = WFmtX1(S1)
PushIAy WFmtX, WFmtX2(Nm1, Nm2, Len(S1(0)), Sep)
Dim J%
For J = 0 To UB(A)
    PushI WFmtX, S1(J) & Sep & S2(J)
Next
End Function

Private Function WFmtX1$(A$())
Dim J%
For J = 0 To UB(A)
    If HasSpc(A(J)) Then WFmtX1 = " | ": Exit Function
Next
WFmtX1 = " "
End Function

Private Function WFmtX2(Nm1$, Nm2$, W%, Sep$) As String()
If Nm1 = "" And Nm2 = "" Then Exit Function
PushI WFmtX2, XAlignL(Nm1, W) & Sep & Nm2
End Function

Private Function Any_Lines(A() As S1S2) As Boolean
Dim J&
Any_Lines = True
For J = 0 To UB(A)
    With A(J)
        If IsLines(.S1) Then Exit Function
        If IsLines(.S2) Then Exit Function
    End With
Next
Any_Lines = False
End Function

Private Function ZW%(LinesAy$(), Nm$)
ZW = Max(LinesAy_Wdt(LinesAy), Len(Nm))
End Function

Private Function ZW1%(A() As S1S2, Nm1$)
ZW1 = ZW(S1S2Ay_Sy1(A), Nm1)
End Function

Private Function ZW2%(A() As S1S2, Nm2$)
ZW2 = ZW(S1S2Ay_Sy2(A), Nm2)
End Function

Property Get Samp_S1S2Ay() As S1S2()
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushObj Samp_S1S2Ay, S1S2(XRpl_VBar(A1(J)), XRpl_VBar(A2(J)))
Next
End Property

Property Get Samp_S1S2Ay1() As S1S2()
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushObj Samp_S1S2Ay1, S1S2(A1(J), A2(J))
Next
End Property

Private Sub Z_S1S2Ay_Fmt()
Dim A() As S1S2, Nm1$, Nm2$
'Nm1 = "AA": Nm2 = "BB": A = Samp_S1S2Ay: GoSub Tst
'Nm1 = "":   Nm2 = "":   A = Samp_S1S2Ay: GoSub Tst
Nm1 = "AA":   Nm2 = "BB":   A = Samp_S1S2Ay1: GoSub Tst
Exit Sub
Tst:
    Act = S1S2Ay_Fmt(A, Nm1, Nm2)
    Brw Act
    Return
End Sub

Private Sub Z()
Z_S1S2Ay_Fmt
MVb_X_S1S2_Fmt:
End Sub
