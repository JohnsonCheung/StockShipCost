Attribute VB_Name = "MVb_Str_Bkt"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb__Bkt."

Private Sub Z_Str_BktPos_ASG()
Dim A$, EptFmPos%, EptToPos%
'
A = "(A(B)A)A"
EptFmPos = 1
EptToPos = 7
GoSub Tst
'
A = " (A(B)A)A"
EptFmPos = 2
EptToPos = 8
GoSub Tst
'
A = " (A(B)A )A"
EptFmPos = 2
EptToPos = 9
GoSub Tst
'
Exit Sub
Tst:
    Dim ActFmPos%, ActToPos%
    Str_BktPos_ASG A, "(", ActFmPos, ActToPos
    Ass IsEq(ActFmPos, EptFmPos)
    Ass IsEq(ActToPos, EptToPos)
    Return
End Sub

Private Sub Z_XBrk_Bkt()
Dim A$, OpnBkt$
A = "aaaa((a),(b))xxx":    OpnBkt = "(":          Ept = Ap_Sy("aaaa", "(a),(b)", "xxx"): GoSub Tst
Exit Sub
Tst:
    Act = XBrk_Bkt(A, OpnBkt)
    C
    Return
End Sub
Function Str_IsEq(A, B, Optional CmpMth As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
Str_IsEq = StrComp(A, B, CmpMth) = 0
End Function

Sub Str_BktPos_ASG(A, OpnBkt$, OFmPos%, OToPos%)
Const CSub$ = CMod & "Str_BktPos_ASG"
OFmPos = 0
OToPos = 0
'-- OFmPos
    Dim Q1$, Q2$
    Q1 = OpnBkt
    Q2 = OpnBkt_ClsBktt(OpnBkt)

    OFmPos = InStr(A, Q1)
    If OFmPos = 0 Then Exit Sub
'-- OToPos
    Dim NOpn%, J%
    For J = OFmPos + 1 To Len(A)
        Select Case Mid(A, J, 1)
        Case Q2
            If NOpn = 0 Then
                OToPos = J
                Exit For
            End If
            NOpn = NOpn - 1
        Case Q1
            NOpn = NOpn + 1
        End Select
    Next
    If OToPos = 0 Then XThw CSub, "The bracket-[Q1]-[Q2] in [Str] has is not in pair: [Q1-Pos] is found, but not Q2-pos is 0", Q1, Q2, A, OFmPos
End Sub

Function OpnBkt_ClsBktt$(OpnBkt$)
Select Case OpnBkt
Case "(": OpnBkt_ClsBktt = ")"
Case "[": OpnBkt_ClsBktt = "]"
Case "{": OpnBkt_ClsBktt = "}"
Case Else: Stop
End Select
End Function

Function XBrk_Bkt(A, Optional OpnBkt$ = vbOpnBkt) As String()
Dim P1%, P2%
    Str_BktPos_ASG A, OpnBkt, _
    P1, P2
Dim A1$, A2$, A3$
A1 = Left(A, P1 - 1)
A2 = Mid(A, P1 + 1, P2 - P1 - 1)
A3 = Mid(A, P2 + 1)
XBrk_Bkt = Ap_Sy(A1, A2, A3)
End Function

Function XTak_BetBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
Str_BktPos_ASG A, OpnBkt, P1, P2
XTak_BetBkt = Mid(A, P1 + 1, P2 - P1 - 1)
End Function

Function XTak_AftBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
   Str_BktPos_ASG A, OpnBkt, P1, P2
If P2 = 0 Then Exit Function
XTak_AftBkt = Mid(A, P2 + 1)
End Function

Function XTak_BefBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
   Str_BktPos_ASG A, OpnBkt, P1, P2
If P1 = 0 Then Exit Function
XTak_BefBkt = Left(A, P1 - 1)
End Function




Private Sub Z()
Z_XBrk_Bkt
Z_Str_BktPos_ASG
MVb_Str_Bkt:
End Sub
