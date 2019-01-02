Attribute VB_Name = "MVb_Lin_Vbl"
Option Compare Binary
Option Explicit
Function Vbl_LasLin$(Vbl)
Vbl_LasLin = Ay_LasEle(SplitVBar(Vbl))
End Function

Function VblAy_IsVdt(A$()) As Boolean
If Sz(A) = 0 Then VblAy_IsVdt = True: Exit Function
Dim I
For Each I In A
    If Not IsVdtVbl(CStr(I)) Then Exit Function
Next
VblAy_IsVdt = True
End Function

Function VblAy_Wdt%(VblAy$())
Dim W%(), J%
For J = 0 To UB(VblAy)
    Push W, Ay_Wdt(VblLy(VblAy(J)))
Next
VblAy_Wdt = AyMax(W)
End Function

Function VblDic(Vbl, Optional JnSep$ = vbCrLf) As Dictionary
Set VblDic = New_Dic_LY(SplitVBar(Vbl), JnSep)
End Function

Function VblLasLin$(Vbl)
VblLasLin = Ay_LasEle(SplitVBar(Vbl))
End Function
Function VblLinesOPT$(Vbl, Optional Pfx$, Optional Ident0%, Optional Sfx$, Optional Wdt0%)
VblLinesOPT = JnCrLf(VblFmt(VblLines(Vbl), Pfx, Ident0, Sfx, Wdt0))
End Function

Function VblLines$(Vbl)
VblLines = JnCrLf(VblLy(Vbl))
End Function
Function VblLy(Vbl) As String()
VblLy = SplitVBar(Vbl)
End Function
Function VblFmt(A, Optional Pfx$, Optional Ident0%, Optional Sfx$, Optional Wdt0%) As String()

'Function VblAlignAsLy(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%) As String()
'Ass IsVbl(Vbl)
'If IsEmp(Vbl) Then Exit Function
'Dim Wdt%
'    Dim W%
'    W = VblWdt(Vbl)
'    If W > WdtOpt Then
'        Wdt = W
'    Else
'        Wdt = WdtOpt
'    End If
'Dim Ident%
'    If Ident < 0 Then
'        Ident = 0
'    Else
'        Ident = IdentOpt
'    End If
'Dim O$()
'    Dim Ay$()
'    Ay = SplitVBar(Vbl)
'    Dim J%, A$, U&, S$, S1$, P$
'    U = UB(Ay)
'    P = IIf(Pfx <> "", Pfx & " ", "")
'    S1 = Space(Ident)
'    For J = 0 To U
'        If J = 0 Then
''            S = AlignL(P, Ident, DoNotCut:=True)
'        Else
'            S = S1
'        End If
''        A = S & AlignL(Ay(J), Wdt, ErIfNotEnoughWdt:=True)
'        If J = U Then
'            A = A & " " & Sfx
'        End If
'        Push O, A
'    Next
'VblAlignAsLy = O
'End Function







Dim Ly$()
Ly = VblLy(A)
Dim Wdt%
    Wdt = Ay_Wdt(Ly)
    If Wdt < Wdt0 Then
        Wdt = Wdt0
    End If
Dim Ident%
    If Ident < 0 Then
        Ident = 0
    Else
        Ident = Ident0
    End If
    If Pfx <> "" Then
        If Ident < Len(Pfx) Then
            Ident = Len(Pfx) + 1
        End If
    End If
Dim O$()
    Dim Ay$()
    Ay = SplitVBar(A)
    Dim J%, U&, S$, S1$, P$
    U = UB(Ay)
    P = IIf(Pfx <> "", Pfx & " ", "")
    S1 = Space(Ident)
    For J = 0 To U
        If J = 0 Then
'            S = AlignL(P, Ident, DoNotCut:=True)
        Else
            S = S1
        End If
'        A = S & AlignL(Ay(J), Wdt, ErIfNotEnoughWdt:=True)
        If J = U Then
            A = A & " " & Sfx
        End If
        PushI O, A
    Next
VblFmt = O
End Function

Function VblLyDry(A$()) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O()
   Dim I
   For Each I In A
       Push O, AyTrim(SplitVBar(CStr(I)))
   Next
VblLyDry = O
End Function

Private Sub Z_VblLyDry()
Dim VblLy$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
GoSub Tst
Exit Sub
Tst:
    Act = VblLyDry(VblLy)
    Ass DryIsEq(CvAy(Act), CvAy(Ept))
    Return
End Sub


Private Sub ZZ_VblFmt()
Ay_XDmp VblFmt("lksfj|lksdfjldf|lskdlksdflsdf|sdkjf", "Select")
End Sub

Private Sub ZZ_VblLyDry()
Dim VblLy$()
Dim Exp$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
Dim Act()
Act = VblLyDry(VblLy)
End Sub


Private Sub Z()
Z_VblLyDry
MVb_Lin_Vbl:
End Sub
