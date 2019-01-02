Attribute VB_Name = "MXls_Z_Lo_Fmt_App"
Option Compare Binary
Option Explicit
Dim Fny$(), Lo As ListObject, LoNm$

Sub Lo_Fmt_APP(A As ListObject)
Fny = Lo_Fny(A)
LoNm = A.Name
Set Lo = A
XSet_Wdt Lo_Fmt_APP1("Select Wdt,FldLikss from Lo_FmtWdt Where LoNm='?'")
XSet_Lvl Lo_Fmt_APP1("Select Wdt,FldLikss from Lo_FmtWdt Where LoNm='?'")
XSet_Align Lo_Fmt_APP1("Select Wdt,FldLikss from Lo_FmtWdt Where LoNm='?'")
XSet_Bdr Lo_Fmt_APP1("")
XSet_Bet Lo_Fmt_APP1("")
XSet_Cor Lo_Fmt_APP1("")
XSet_Fml Lo_Fmt_APP1("")
XSet_Fmt Lo_Fmt_APP1("")
XSet_Lbl Lo_Fmt_APP1("")
XSet_Lvl Lo_Fmt_APP1("")
XSet_Tit Lo_Fmt_APP1("")
XSet_Tot Lo_Fmt_APP1("")
XSet_Wdt Lo_Fmt_APP1("")
End Sub
Private Function Lo_Fmt_APP1(QQ$) As Drs
Set Lo_Fmt_APP1 = QQ_Drs(QQ, LoNm)
End Function
Private Sub XSet_Align(A As Drs)
Dim Dry()
Dry = A.Dry
XX A, "HAlign FldLikss"
XSet_Align2 XSet_Align1(Dry, "Left"), xlHAlignLeft
XSet_Align2 XSet_Align1(Dry, "Right"), xlHAlignRight
XSet_Align2 XSet_Align1(Dry, "Center"), xlHAlignCenter
End Sub

Private Function XSet_Align1$(Dry(), HAlign$)
Dim A$()
A = DrySy(Dry_XWh_ColEq(Dry, 0, HAlign), 1)
Select Case Sz(A)
Case 0
Case 1: XSet_Align1 = A(0)
Case Else: Stop
End Select
End Function

Private Sub XSet_Align2(Fldss$, A As XlHAlign)
Dim F
For Each F In Fny
    If XLik_Likss(F, Fldss) Then Lc_XSet_Align Lo, F, A
Next
End Sub

Private Sub XSet_Bdr(A As Drs)
XX A, ""
Dim L$(), R$(), C$(), Dry()
Dry = A.Dry
L = XSet_Bdr1(Dry, "Left")
R = XSet_Bdr1(Dry, "Right")
C = XSet_Bdr1(Dry, "Center")
XSet_BdrLeft L
XSet_BdrLeft C
XSet_BdrRight C
XSet_BdrRight R
End Sub

Private Function XSet_Bdr1(Dry(), HBdr$) As String()

End Function

Private Function XSet_BdrLeft(FldLikAy$())
Dim F
For Each F In Fny
    If XLik_LikAy(F, FldLikAy) Then Lc_XSet_BdrLeft Lo, F
Next
End Function

Private Function XSet_BdrRight(FldLikAy$())
Dim F
For Each F In Fny
    If XLik_LikAy(F, FldLikAy) Then Lc_XSet_BdrRight Lo, F
Next
End Function

Private Sub XSet_Bet(A As Drs)
XX A, ""
Dim Dr, C$, X$, Y$
For Each Dr In AyNz(A.Dry)
    AyAsg Dr, C, X, Y
    Lo.ListColumns(C).DataBodyRange.Formula = QQ_Fmt("=Sum([?]:[?])", X, Y)
Next
End Sub

Private Sub XSet_Cor(A As Drs)
XX A, "Colr FldLikss"
Dim Dr
For Each Dr In AyNz(A.Dry)
    XSet_Cor1 Dr
Next
End Sub

Private Sub XSet_Cor1(Dr)
Dim Colr&, FldLikss$, F
AyAsg Dr, Colr, FldLikss
For Each F In Fny
    If XLik_Likss(F, FldLikss) Then Lc_XSet_Cor Lo, F, Colr
Next
End Sub

Private Sub XSet_Fml(A As Drs)
Dim C$, Dr, Fml$
For Each Dr In AyNz(A.Dry)
    AyAsg Dr, C, Fml
    Lc_XSet_Fml Lo, C, Fml
Next
End Sub

Private Sub XSet_Fmt(A As Drs)
Dim Dr
For Each Dr In AyNz(A.Dry)
    XSet_Fmt1 Dr
Next
End Sub

Private Sub XSet_Fmt1(Dr)
Dim F
Dim Fmt$, FldLikss$
For Each F In Fny
    AyAsg Dr, Fmt, FldLikss
    If XLik_Likss(F, FldLikss) Then
        Lc_XSet_Fmt Lo, F, Fmt
        Exit Sub
    End If
Next
End Sub

Private Sub XSet_Lbl(A As Drs)
XX A, "Fld Lbl"
Dim Dr
For Each Dr In AyNz(A.Dry)
    XSet_Lbl1 Dr
Next
End Sub

Private Sub XSet_Lbl1(Dr)
Dim F$, Lbl$
AyAsg Dr, F, Lbl
If Not Ay_XHas(Fny, F) Then Exit Sub
End Sub

Private Sub XSet_Lvl(A As Drs)
Dim Dr
XX A, "Lvl FldLikss"
For Each Dr In AyNz(A.Dry)
    XSet_Lvl1 Dr
Next
End Sub

Private Sub XSet_Lvl1(Dr)
Dim F
Dim Lvl As Byte, FldLikss$
AyAsg Dr, Lvl, FldLikss
For Each F In Fny
    If XLik_Likss(F, FldLikss) Then
        Lc_XSet_Lvl Lo, F, Lvl
    End If
Next
End Sub

Private Sub XSet_Tit(A As Drs)
XX A, "Fld Tit"
'Lo_XSet_Tit Lo, Tit
End Sub

Private Sub XSet_Tot(A As Drs)
'XSet_Tot1 Ay_FstRmvT1(Tot, "Sum"), xlTotalsCalculationSum
'XSet_Tot1 Ay_FstRmvT1(Tot, "Cnt"), xlTotalsCalculationCount
'XSet_Tot1 Ay_FstRmvT1(Tot, "Avg"), xlTotalsCalculationAverage
End Sub

Private Sub XSet_Tot1(FldLikss$, B As XlTotalsCalculation)
Dim F
For Each F In Fny
    If XLik_Likss(F, FldLikss) Then Lc_XSet_Tot Lo, F, B
Next
End Sub

Private Sub XSet_Wdt(A As Drs)
XX A, "Wdt FldLikss"
Dim Dr
For Each Dr In AyNz(A.Dry)
    XSet_Wdt1 Dr
Next
End Sub

Private Sub XSet_Wdt1(Dr) _
' Dr is Array-of-W-FldLikss
Dim W%, FldLikss$
AyAsg Dr, W, FldLikss
Dim F
For Each F In Fny
    If XLik_Likss(F, FldLikss) Then Lc_XSet_Wdt Lo, F, W
Next
End Sub

Private Sub XX(A As Drs, Ssl$)
Ay_IsEq_XAss A, Ssl_Sy(Ssl)
End Sub
