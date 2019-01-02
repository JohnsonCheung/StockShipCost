Attribute VB_Name = "MIde_Mth_Lin_Sig_Pm"
Option Compare Binary
Option Explicit

Function Lin_NPrm(A) As Byte
Lin_NPrm = SubStrCnt(XTak_BetBkt(A), ",")
End Function

Function Lin_PmNy(A) As String()
Dim L$
L = XRmv_Mdy(A)
If XShf_MthShtTy(L) = "" Then Exit Function
L = XTak_BetBkt(L)
Lin_PmNy = SplitComma(L)
End Function

Function MthLinPmAy(MthLin$) As MthPm()
Dim ArgStr$
    ArgStr = XTak_BetBkt(MthLin)
Dim P$()
    P = SplitComma(ArgStr)
Dim O() As MthPm
    Dim U%: U = UB(P)
    ReDim O(U)
    Dim J%
    For J = 0 To U
        'O(J) = MthPm(P(J))
    Next
MthLinPmAy = O
End Function

Function MthPm(MthPmStr$) As MthPm
Dim L$: L = MthPmStr
Const Opt$ = "Optional"
Const PA$ = "ParamArray"
Dim TyChr$
With MthPm
    .IsOpt = XShf_PfxSpc(L, Opt)
    .IsPmAy = XShf_PfxSpc(L, PA)
    .Nm = XShf_Nm(L)
    .Ty.TyChr = XShf_Chr(L, "!@#$%^&")
    .Ty.IsAy = XShf_Pfx(L, "()") = "()"
End With
End Function

Function MthPmAyNy(A() As MthPm) As String()
Dim J%, O$()
For J = 0 To MthPmUB(A)
    Push O, A(J).Nm
Next
MthPmAyNy = O
End Function

Function MthPmSz&(A() As MthPm)
On Error Resume Next
MthPmSz = UBound(A) + 1
End Function

Function MthPmTyAsTyNm$(A As MthPmTy)
With A
    If .TyChr <> "" Then MthPmTyAsTyNm = MthTyChrAsTyStr(.TyChr): Exit Function
    If .TyAsNm = "" Then
        MthPmTyAsTyNm = "Variant"
    Else
        MthPmTyAsTyNm = .TyAsNm
    End If
End With
End Function

Function MthPmTyShtNm$(RetTy As MthPmTy)
Dim Ay$
Dim O$
    With RetTy
        If .IsAy Then Ay = "Ay"
        Select Case .TyChr
        Case "!": O = "Sng"
        Case "@": O = "Cur"
        Case "#": O = "Dbl"
        Case "$": O = "Str"
        Case "%": O = "Int"
        Case "^": O = "LngLng"
        Case "&": O = "Lng"
        End Select
        If O = "" Then
            O = .TyAsNm
        End If
        If O = "" Then
            O = "Var"
        End If
    End With
    Select Case O
    Case "String": O = "Str"
    Case "Integer": O = "Int"
    Case "Long": O = "Lng"
    Case "Currency": O = "Cur"
    Case "Single": O = "Sng"
    Case "Double": O = "Dbl"
    Case "LongLong": O = "Lng"
    End Select
    O = O & Ay
    If O = "StrAy" Then O = "Sy"
MthPmTyShtNm = O
End Function

Function MthPmUB&(A() As MthPm)
MthPmUB = MthPmSz(A) - 1
End Function

Function PushMthPm(O() As MthPm, M As MthPm) As MthPm
Dim N&: N = MthPmSz(O)
ReDim Preserve O(N)
O(N) = M
End Function

Private Sub Z()
Exit Sub
'Lin_NPrm
'Lin_PmNy A
'MthLinArgStr
'MthLinPmAy
'MthPm
'MthPmAyNy
'MthPmSz
'MthPmTyAsTyNm
'MthPmTyShtNm
'MthPmUB
'PushMthPm
End Sub
