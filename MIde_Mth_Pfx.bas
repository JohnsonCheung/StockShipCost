Attribute VB_Name = "MIde_Mth_Pfx"
Option Compare Binary
Option Explicit

Private Sub Z_MthPfx()
Ass MthPfx("Add_Cls") = "Add"
End Sub

Private Sub ZZ_MthPfx()
Dim Ay$(): Ay = Vbe_MthNy(CurVbe)
Dim Ay1$(): Ay1 = AyMap_Sy(Ay, "MthPfx")
Ws_XVis AyAB_Ws_Hori(Ay, Ay1)
End Sub

Function MthPfx$(MthNm)
Dim A0$
    A0 = Brk1(XRmv_PfxAy(MthNm, SplitVBar("ZZ_|Z_")), "__").S1
With Brk2(A0, "_")
    If .S1 <> "" Then
        MthPfx = .S1
        Exit Function
    End If
End With
Dim P2%
Dim Fnd As Boolean
    Dim C%
    Fnd = False
    For P2 = 2 To Len(A0)
        C = Asc(Mid(A0, P2, 1))
        If Asc_IsLCase(C) Then Fnd = True: Exit For
    Next
'---
    If Not Fnd Then Exit Function
Dim P3%
Fnd = False
    For P3 = P2 + 1 To Len(A0)
        C = Asc(Mid(A0, P3, 1))
        If Asc_IsUCase(C) Or Asc_IsDigit(C) Then Fnd = True: Exit For
    Next
'--
If Fnd Then
    MthPfx = Left(A0, P3 - 1)
    Exit Function
End If
MthPfx = MthNm
End Function


Private Sub Z()
Z_MthPfx
MIde_Mth_Pfx:
End Sub
