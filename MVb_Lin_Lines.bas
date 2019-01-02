Attribute VB_Name = "MVb_Lin_Lines"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Lin_Lines."
Private Sub Z_Lines_XWrap()
Dim A$, W%
A = "lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf sdklf sdklfj dsfj "
W = 80
Ept = Ap_Sy("lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf ", _
"klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf", _
"sdklf sdklfj dsfj ")
GoSub Tst
Exit Sub
Tst:
    Act = Lines_XWrap(A, W)
    C
    Return
End Sub
Function Lines_XWrap(A, Optional Wdt% = 80) As String()
Dim L$, W%, O$, J%
W = Wdt
If W < 10 Then W = 10: XDmp_Ly CSub, "Given Wdt is too small, 10 is used", "Wdt Lines", Wdt, A
L = A
While Len(L) > 0
    J = J + 1: If J >= 1000 Then XThw CSub, "Program error: Given Lines is too line", "Lines Wdt", A, Wdt
    O = Left(L, W)
    L = Mid(L, W + 1)
    PushI Lines_XWrap, O
Wend
End Function

Private Sub ZZ_LinesAy_LyPad()
Dim A$()
Push A, XRpl_VBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf")
Push A, XRpl_VBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf|sdf")
Push A, XRpl_VBar("ksdjlfdf|sdklfjdsfdf|skldfjdf|lskdf|slkdjf|sdlf||")
Push A, XRpl_VBar("ksdjlfdf|sdklfjsdfdsfdsf|skldsdffjdf")
D LinesAy_LyPad(A)
End Sub


Private Sub ZZ_Lines_XTrim_End()
Dim Lines$: Lines = XRpl_VBar("lksdf|lsdfj|||")
Dim Act$: Act = Lines_XTrim_End(Lines)
Debug.Print Act & "<"
Stop
End Sub

Private Sub ZZ_Lines_LasNLin()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
'Debug.Print fLasN(A, 3)
End Sub

Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function

Function LinesAy_LyPad(A$()) As String()
LinesAy_LyPad = LyPad(LinesAy_Ly(A))
End Function

Function LinesAy_Ly(A) As String()
Dim Lines
For Each Lines In AyNz(A)
    PushIAy LinesAy_Ly, SplitLines(Lines)
Next
End Function

Function LinesAy_Lines$(A$())
LinesAy_Lines = JnCrLf(LinesAy_Ly(A))
End Function

Function LinesAy_Wdt%(A$())
Dim O%, L
For Each L In AyNz(A)
   O = Max(O, Lines_Wdt(L))
Next
LinesAy_Wdt = O
End Function

Function Lines_BoxLy(A$) As String()
Lines_BoxLy = LyBoxLy(SplitCrLf(A))
End Function

Sub LinesBrkAsg(A$, Ny0, ParamArray OLyAp())
Dim Ny$(), L, T1$, T2$, NmDic As Dictionary
Ny = CvNy(Ny0)
Set NmDic = AyIxDic(Ny)
For Each L In SplitCrLf(A)
    Select Case XTak_FstChr(L)
    Case "'", " "
    Case Else
        BrkAsg L, " ", T1, T2
        If NmDic.Exists(T1) Then
            Push OLyAp(NmDic(T1)), T2 '<----
        End If
    End Select
Next
End Sub


Private Sub Z_Lines_XTrim_End()
Dim Lines$: Lines = XRpl_VBar("lksdf|lsdfj|||")
Dim Act$: Act = Lines_XTrim_End(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function Lines_LasNLin$(A, N%)
Lines_LasNLin = JnCrLf(Ay_XWh_LasN(SplitCrLf(A), N))
End Function

Function Lines_LasLin$(A)
If A = "" Then Exit Function
Lines_LasLin = Ay_LasEle(Lines_Ly(A))
End Function

Function Lines_LinCnt&(A)
Lines_LinCnt = Sz(SplitCrLf(A))
End Function

Function Lines_Ly(A) As String()
Lines_Ly = SplitLines(A)
End Function

Function LinesSplit(A) As String()
LinesSplit = SplitCrLf(A)
End Function

Function Lines_SqH(A) As Variant()
Lines_SqH = AySqH(Lines_Ly(A))
End Function

Function Lines_SqV(A) As Variant()
Lines_SqV = AySqV(Lines_Ly(A))
End Function

Function Lines_XTrim_End$(A$)
Lines_XTrim_End = JnCrLf(Ly_XTrim_End(SplitCrLf(A)))
End Function

Function Lines_XIdent$(Lines$, Optional Space% = 4)
Dim O$(), S$, L
S = VBA.Space(Space)
For Each L In AyNz(SplitCrLf(Lines))
    PushI O, S & L
Next
Lines_XIdent = JnCrLf(O)
End Function

Function Lines_Vbl$(A$)
Const CSub$ = CMod & "Lines_Vbl"
If HasSubStr(A, "|") Then XThw CSub, "Given [Lines] has |", A
Lines_Vbl = Replace(A, vbCrLf, "|")
End Function

Function Lines_Wdt%(A)
Lines_Wdt = Ay_Wdt(SplitLines(A))
End Function


Private Sub Z()
Z_Lines_XTrim_End
Z_Lines_XWrap
MVb_Lin_Lines:
End Sub
