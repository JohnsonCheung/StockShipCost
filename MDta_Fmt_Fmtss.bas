Attribute VB_Name = "MDta_Fmt_Fmtss"
Option Compare Binary
Option Explicit

Function Dry_Fmtss(A()) As String()
Dim W%(), Dr, O$()
W = Dry_WdtAy(A)
For Each Dr In AyNz(A)
    PushI O, Dr_Fmt(Dr, W, " ")
Next
Dry_Fmtss = O
End Function
Function Dr_StrDr_UsingWdtAy(A, WdtAy%()) As String()
Dim UDr%
   UDr = UB(A)
Dim O$()
   ReDim O(UB(WdtAy))
   Dim W, J%
   J = 0
   For Each W In WdtAy
        If UDr >= J Then
            O(J) = XAlignL(A(J), W)
        Else
            O(J) = Space(W)
        End If
        J = J + 1
   Next
Dr_StrDr_UsingWdtAy = O
End Function
Function Dr_Fmt$(A, WdtAy%(), Optional HidGrid As Boolean)
Dim Sep$, Pfx$, Sfx$
If HidGrid Then
    Sep = " "
Else
    Sep = " | "
    Pfx = "| "
    Sfx = " |"
End If
Dr_Fmt = Pfx & Jn(Dr_StrDr_UsingWdtAy(A, WdtAy), Sep) & Sfx
End Function

Function Ay_Dry_BySepSS(A, SepSS$) As Variant()
Dim Lin, SepAy$()
SepAy = Ssl_Sy(SepSS)
For Each Lin In AyNz(A)
    PushI Ay_Dry_BySepSS, Lin_Dr_BySepAy(Lin, SepAy)
Next
End Function

Sub Drs_FmtssDmp(A As Drs)
D Drs_Fmtss(A)
End Sub

Function Drs_Fmtss(A As Drs) As String()
Drs_Fmtss = Dry_Fmtss(CvAy(ItmAddAy(A.Fny, A.Dry)))
End Function

Sub Drs_FmtssBrw(A As Drs)
Brw Drs_Fmtss(A)
End Sub

Function Dry_FmtssCell(A()) As Variant()
Dim Dr
For Each Dr In AyNz(A)
Stop
'    Push Dry_FmtssCell, Dr_FmtCell(Dr) ' Fmtss(X)
Next
End Function
Private Sub Z_Lin_Dr_BySepAy()
Dim Lin$, SepAy$()
SepAy = Ssl_Sy(". . . . . .")
Lin = "AStkShpCst_Rpt.OupFx.Fun."
Ept = Ap_Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
GoSub Tst
Tst:
    Act = Lin_Dr_BySepAy(Lin, SepAy)
    C
End Sub

Function Lin_Dr_BySepAy(Lin, SepAy$()) As String()
Dim Sep, P%, L$, J%, W%
L = Lin
W = 1
For Each Sep In SepAy
    P = InStr(W, L, Sep)
    If P = 0 Then Exit For
    Push Lin_Dr_BySepAy, Left(L, P - 1)
    L = Mid(L, P)
    J = J + 1
    W = Len(SepAy(J)) + 1
Next
If L <> "" Then
    PushI Lin_Dr_BySepAy, L
End If
End Function
Function DotAy_Dry(A) As Variant()
Dim L
For Each L In AyNz(A)
    PushI DotAy_Dry, SplitDot(L)
Next
End Function

Function DotAy_Fmt(A) As String()
DotAy_Fmt = Dry_Fmt(DotAy_Dry(A), HidGrid:=True)
End Function

Function Ay_Fmt(A, SepSS$) As String()
Ay_Fmt = Dry_Fmtss(Ay_Dry_BySepSS(A, SepSS))
End Function
