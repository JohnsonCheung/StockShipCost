Attribute VB_Name = "MIde_Gen_Const_MthLines"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Gen_Const_ConstValMthLines."

Function ConstVal_PrpLines$(ConstVal$, Nm$, IsPub As Boolean) _
' Return [MthLines] by [ConstVal$] and [Nm$]
Const CSub$ = CMod & "ConstValMthLines"
If ConstVal = "" Then XThw_Msg CSub, "Given ConstVal is blank"
Dim A$()
Dim NChunk%
    A = SplitCrLf(ConstVal)
    NChunk = WNChunk(Sz(A))
Dim O$()
    Dim J%
    For J = 0 To NChunk - 1
        PushI O, WChunk(A, J)
    Next
    PushI O, WLasLin(Nm, NChunk)
ConstVal_PrpLines = WMthLines(JnCrLf(O), Nm, IsPub)
End Function

Private Function WChunk$(ConstLy$(), IChunk%)
If Sz(ConstLy) = 0 Then Stop
Dim Ly$()
    Ly = AyMid(ConstLy, IChunk * 20, 20)
Dim O$()
    Dim L$, J&, U&
    U = UB(Ly)
    For J = 0 To U
        L = XQuote_Dbl(Ly(J))
        Select Case True
        Case J = 0 And J = U: Push O, QQ_Fmt("Const A_?$ = ?", IChunk + 1, L)
        Case J = 0:           Push O, QQ_Fmt("Const A_?$ = ? & _", IChunk + 1, L)
        Case J = U:           Push O, "vbCrLf & " & L
        Case Else:            Push O, "vbCrLf & " & L & " & _"
        End Select
    Next
WChunk = JnCrLf(O) & vbCrLf
End Function

Private Function WLasLin$(Nm$, NChunk%)
Dim B$
    Dim O$(), J%
    For J = 1 To NChunk
        PushI O, "A_" & J
    Next
    B = Join(O, " & vbCrLf & ")
WLasLin = Nm & " = " & B
End Function

Private Function WMthLines$(Lines$, Nm$, IsPub As Boolean)
Dim L1$, L2$
L1 = IIf(IsPub, "", "Private ") & "Function " & Nm & "$()" & vbCrLf
L2 = vbCrLf & "End Function"
WMthLines = vbCrLf & L1 & Lines & L2
End Function

Private Function WNChunk%(Sz%)
WNChunk = ((Sz - 1) \ 20) + 1
End Function

Private Sub Z_ConstVal_PrpLines()
Const CSub$ = CMod & "Z_ConstValMthLines"
'GoSub Cas_Simple
GoSub Cas_Complex
'GoSub Cas_Complex1
Exit Sub
'--
Dim Nm$, ConstVal$, IsPub As Boolean
Dim IsEdt As Boolean, CAs$
Cas_Complex1:
    CAs = "Complex1"
    IsEdt = False
    Nm = "ZZ_B"
    ConstVal = TstTxt(CurPj_Nm, CSub, CAs, "ConstVal", IsEdt)
    Ept = TstTxt(CurPj_Nm, CSub, "Complex1", "Ept", IsEdt)
    IsPub = True
    GoTo Tst

Cas_Complex:
    IsEdt = True
    ConstVal = MdMthNm_Lines(CurMd, "WChunk")
    StrBrw ConstVal
    Stop
    Nm = "ZZ_A"
    IsPub = True
    Ept = TstTxt(CurPj_Nm, CSub, "Complex", "Ept", IsEdt)
    GoTo Tst
'
Cas_Simple:
    IsEdt = False
    Nm = "ZZ_A"
    ConstVal = "AAA"
    Ept = JnCrLf(Array("", _
        "Private Function ZZ_A$()", _
        "Const A_1$ = ""AAA""", _
        "", _
        "ZZ_A = A_1", _
        "End Function"))
    GoTo Tst
Tst:
    If IsEdt Then Return
    If ConstVal = "" Then Stop
    Act = ConstVal_PrpLines(ConstVal, Nm, IsPub)
    'Brw Act: Stop
    C
    TstOk CSub, CAs
    Return
End Sub

Private Sub ZZ()
Dim A$
Dim B As Boolean
ConstVal_PrpLines A, A, B
End Sub

Private Sub Z()
Z_ConstVal_PrpLines
End Sub
