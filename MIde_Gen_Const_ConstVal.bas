Attribute VB_Name = "MIde_Gen_Const_ConstVal"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Gen_Const_MthLines_ConstVal."

Private Property Get AA$()
Const A_1$ = "Johnsonlskf;lskf" & _
vbCrLf & "ksjdflkjs dflkj" & _
vbCrLf & "lksdjf lskdfj" & _
vbCrLf & ""

AA = A_1
End Property

Function MthLines_ConstVal$(MthLines)
'Return a string constant from the source code.  A reverse of [ConstValMthLines]
Dim O$, ConstLines
For Each ConstLines In AyNz(MthLines_ConstLinesAy(MthLines))
    O = O & ConstLines_ConstVal(ConstLines)
Next
MthLines_ConstVal = O
End Function

Private Function ConstLines_ConstVal$(C)
Dim I, O$(), A$, B$
For Each I In SplitCrLf(C)
    A = TakBetFstLas(I, """", """")
    B = Replace(A, """""", """")
    PushI O, B
Next
ConstLines_ConstVal = JnCrLf(O)
End Function

Private Function MthLines_ConstLinesAy(MthLines) As String()
Dim Ay$(), O$
O = MthLines
Lp:
    Ay = TakP123(O, "Const", vbCrLf & vbCrLf)
    If Sz(Ay) = 3 Then
        PushI MthLines_ConstLinesAy, Ay(1)
        O = Ay(2)
        GoTo Lp
    End If
End Function

Private Sub Z_MthLines_ConstVal()
Const CSub$ = CMod & "Z_MthLines_ConstVal"
Dim IsEdt As Boolean, MthLines$, CAs$
GoSub Cas_Complex
GoSub Cas_Simple
Exit Sub
Cas_Complex:
    IsEdt = False
    CAs = "Complex"
    MthLines = TstTxt(CurPj_Nm, CSub, CAs, "MthLines", IsEdt)
    Ept = TstTxt(CurPj_Nm, CSub, CAs, "Ept", IsEdt)
    If IsEdt Then Return
    GoTo Tst
Cas_Simple:
    
    Return
Tst:
    Act = MthLines_ConstVal(MthLines)
    'Brw Act: Stop
    C
    TstOk CSub, CAs
    Return
End Sub

Private Sub Z()
Z_MthLines_ConstVal
MIde_Gen_Const_MthLines_ConstVal:
End Sub
