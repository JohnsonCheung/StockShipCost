Attribute VB_Name = "MVb_Align"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Align."

Function XAlignL$(A, W)
Dim L%: L = Len(A)
If L >= W Then
    XAlignL = A
Else
    XAlignL = A & Space(W - Len(A))
End If
End Function

Function XAlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    XAlignR = Space(W - L) & S
Else
    XAlignR = S
End If
End Function


Function XAlignL1$(S$, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = CMod & "XAlignL1"
Dim L%: L = Len(S)
If L > W Then
    If ErIFmnotEnoughWdt Then
        Stop
        'XThw CSub, "Len({S)) > {W}", S, W
    End If
    If DoNotCut Then
        XAlignL1 = S
        Exit Function
    End If
End If

If W >= L Then
    XAlignL1 = S & Space(W - L)
    Exit Function
End If
If W > 2 Then
    XAlignL1 = Left(S, W - 2) + ".."
    Exit Function
End If
XAlignL1 = Left(S, W)
End Function

Function LinesAy_XAlign_LasLin(A$()) As String()
Dim W%: W = LinesAy_Wdt(A)
Dim Lines
For Each Lines In AyNz(A)
    PushI LinesAy_XAlign_LasLin, LinesAlignLasLin(Lines, W)
Next
End Function
Function LinesAlignLasLin$(A, W%)
Stop '
End Function
Function LinesAlign$(A, W%)
Stop '
End Function
Function LinesAy_XAlign_(A$()) As String()
Dim W%: W = LinesAy_Wdt(A)
Dim Lines
For Each Lines In AyNz(A)
    PushI LinesAy_XAlign_, LinesAlign(Lines, W)
Next
End Function
Function VblAlign$(A, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%)
VblAlign = JnVBar(VblFmt(A, Pfx, IdentOpt, Sfx, WdtOpt))
End Function

