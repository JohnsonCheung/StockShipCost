Attribute VB_Name = "MVb_Lin_Shf"
Option Compare Binary
Option Explicit

Function XShf_Val(OLin$, Lblss) As Variant()
'Return Ay, which is
'   Same sz as QLbl-term-cnt
'   Either string of boolean
'   Each element is corresponding to QLbl
'Return OLin
'   if the term match, it will removed from OLin
'QLbl term: is in *LLL ?LLL or LLL
'   *LLL is always first at beginning, mean the value OLin has not lbl
'   ?LLL means the value is in OLin is using LLL => it is true,
'   LLL  means the value in OLin is LLL=VVV
'OLin is
'   VVV VVV=LLL [VVV=L L]
Dim L, Ay$()
Ay = Lin_TermAy(OLin)
For Each L In AyNz(CvNy(Lblss))
    PushI XShf_Val, XShf_Val__ITM(Ay, L)
Next
OLin = JnSpc(Ay_XQuote_SqBkt_IfNeed(Ay))
End Function

Private Function XShf_Val__BOOL(OAy$(), Lbl) As Boolean
Dim J%, L$
L = XRmv_FstChr(Lbl)
For J = 0 To UB(OAy)
    If OAy(J) = L Then
        XShf_Val__BOOL = True
        OAy = Ay_XExl_EleAt(OAy, J)
        Exit Function
    End If
Next
End Function

Private Function XShf_Val__ITM(OAy$(), Lbl)
If Sz(OAy) = 0 Then XShf_Val__ITM = "": Exit Function
'Return either Boolean or string
Select Case XTak_FstChr(Lbl)
Case "*": XShf_Val__ITM = OAy(0): OAy = Ay_XExl_FstEle(OAy)
Case "?": XShf_Val__ITM = XShf_Val__BOOL(OAy, Lbl)
Case Else: XShf_Val__ITM = XShf_Val__STR(OAy, Lbl)
End Select
End Function

Private Function XShf_Val__STR$(OAy$(), Lbl)
Dim J%
For J = 0 To UB(OAy)
    With Brk1(OAy(J), "=")
        If .S1 = Lbl Then
            XShf_Val__STR = .S2
            OAy = Ay_XExl_EleAt(OAy, J)
            Exit Function
        End If
    End With
Next
End Function

Function XShf_BktStr$(OLin$)
Dim O$
O = XTak_BetBkt(OLin): If O = "" Then Exit Function
XShf_BktStr = O
OLin = XTak_AftBkt(OLin)
End Function

Function XShf_Chr$(OLin, ChrList$)
Dim F$, P%
F = XTak_FstChr(OLin)
P = InStr(ChrList, F)
If P > 0 Then
    XShf_Chr = Mid(ChrList, P, 1)
    OLin = Mid(OLin, 2)
    Exit Function
End If
End Function

Function XShf_Pfx(OLin, Pfx) As Boolean
If XHas_Pfx(OLin, Pfx) Then
    OLin = XRmv_Pfx(OLin, Pfx)
    XShf_Pfx = True
End If
End Function

Function XShf_PfxSpc(OLin, Pfx) As Boolean
If XHas_PfxSpc(OLin, Pfx) Then
    OLin = Mid(OLin, Len(Pfx) + 2)
    XShf_PfxSpc = True
End If
End Function

Private Sub Z_XShf_Val()
GoSub Cas1
GoSub Cas2
GoSub Cas3
Exit Sub
Dim Lin$, Lblss, EptLin$
Cas1:
    Lin = "1 Req"
    EptLin = ""
    Lblss = "*XX ?Req"
    Ept = Array("1", True)
    GoTo Tst
Cas2:
    Lin = "A B C=123 D=XYZ"
    Lblss = "?B"
    Ept = Array(True)
    EptLin = "A C=123 D=XYZ"
    GoTo Tst
Cas3:
    Lin = "Txt VTxt=XYZ [Dft=A 1] VRul=123 Req"
    Lblss = "*Ty ?Req ?AlwZLen Dft VTxt VRul"
    Ept = Array("Txt", True, False, "A 1", "XYZ", "123")
    EptLin = ""
    GoTo Tst
Tst:
    Act = XShf_Val(Lin, Lblss)
    C
    Ass Lin = EptLin
    Return
End Sub

Private Sub Z_XShf_BktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = XShf_BktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Private Property Get Z_XShf_Pfx()
Dim O$: O = "AA{|}BB "
Ass XShf_Pfx(O, "{|}") = "AA"
Ass O = "BB "
End Property



Private Sub Z()
Z_XShf_BktStr
Z_XShf_Val
MVb_Lin_XShf_:
End Sub
