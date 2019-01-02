Attribute VB_Name = "MIde_Mth_Lin_XXX"
Option Compare Binary
Option Explicit
Function LinMdy$(A)
LinMdy = Ay_FstEqV(MdyAy, Lin_T1(A))
End Function

Private Sub Z_MthLinRetTy()
Dim MthLin$
Dim A As MthPmTy:
MthLin = "Function MthPm(MthPmStr$) As MthPm"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = "MthPm"
Ass A.IsAy = False
Ass A.TyChr = ""

MthLin = "Function MthPm(MthPmStr$) As MthPm()"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = "MthPm"
Ass A.IsAy = True
Ass A.TyChr = ""

MthLin = "Function MthPm$(MthPmStr$)"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = ""
Ass A.IsAy = False
Ass A.TyChr = "$"

MthLin = "Function MthPm(MthPmStr$)"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = ""
Ass A.IsAy = False
Ass A.TyChr = ""
End Sub

Function Lin_MthKd$(A)
Lin_MthKd = TakMthKd(XRmv_Mdy(A))
End Function

Function LinMthTy$(A)
LinMthTy = XTak_PfxAySpc(XRmv_Mdy(A), MthTyAy)
End Function



Private Sub Z_LinMthTy()
Dim O$(), L
For Each L In MdNm_Src("Fct")
    Push O, LinMthTy(L) & "." & L
Next
Brw O
End Sub



Private Sub Z_Lin_MthKd()
Dim A$
Ept = "Property": A = "Private Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = Lin_MthKd(A)
    C
    Return
End Sub
Function MthLinRetTy(MthLin$) As MthPmTy
If Not HasSubStr(MthLin, "(") Then Exit Function
If Not HasSubStr(MthLin, ")") Then Exit Function
Dim TC$: TC = XTak_LasChr(XTak_BefBkt(MthLin))
With MthLinRetTy
    If IsMthTyChr(TC) Then .TyChr = TC: Exit Function
    Dim Aft$: Aft = XTak_AftBkt(MthLin)
        If Aft = "" Then Exit Function
        If Not XHas_Pfx(Aft, " As ") Then Stop
        Aft = XRmv_Pfx(Aft, " As ")
        If XHas_Sfx(Aft, "()") Then
            .IsAy = True
            Aft = RmvSfx(Aft, "()")
        End If
        .TyAsNm = Aft
        Exit Function
End With
End Function


Private Sub Z()
Z_Lin_MthKd
Z_LinMthTy
Z_MthLinRetTy
MIde_MthLin_XXX:
End Sub
