Attribute VB_Name = "MIde_Ens_CSub_Brk_Mth"
Option Compare Binary
Option Explicit
Function Md_CSubBrkMthAy(A As CodeModule) As CSubBrkMth()
If Md_IsNoLin(A) Then Exit Function
Dim Ix() As FTIx, Src$(), I, Nm$
Src = Md_Src(A)
Nm = Md_Nm(A)
Ix = Src_MthFTIxAyIxAy(Src)
For Each I In AyNz(Ix)
    PushObj Md_CSubBrkMthAy, WBrk(Nm, Src, CvFTIx(I))
Next
End Function

Private Function WBrk(MdNm$, Src, MthFTIx As FTIx) As CSubBrkMth
Set WBrk = New CSubBrkMth
Dim IFm&, ITo&, IsUsingCSub As Boolean, MthNm$
    IFm = MthFTIx.FmIx
    ITo = MthFTIx.ToIx
    IsUsingCSub = WIsUsingCSub(Src, IFm, ITo)
    MthNm = Lin_MthNm(Src(IFm))
With WBrk
    .MdNm = MdNm
    .MthNm = MthNm
    .IsUsingCSub = IsUsingCSub
    .NewCSub = WNewCSub(MthNm)
    .NewLno = WNewLno(Src, IFm, ITo)
    .OldLno = WOldLno(Src, IFm, ITo)
    If .OldLno > 0 Then _
    .OldCSub = Src(.OldLno - 1)
    .NeedDlt = WNeedDlt(IsUsingCSub, .NewCSub, .OldCSub)
    .NeedIns = WNeedIns(IsUsingCSub, .NewCSub, .OldCSub)
End With
End Function

Private Function WOldLno&(Src, IFm&, ITo&)
Dim J&
For J = IFm To ITo
    If XHas_Pfx(Src(J), "Const CSub$") Then
        WOldLno = J + 1
        Exit Function
    End If
Next
End Function

Private Function WNewCSub$(MthNm$)
WNewCSub = "Const CSub$ = CMod & """ & MthNm & """"
End Function

Private Function WNewLno&(Src, IFm&, ITo&)
If IFm = ITo Then Exit Function
Dim J&, Fm&
Fm = WNewLno1(Src, IFm, ITo) ' Ix after the Mth_MthLin line
For J = Fm To ITo
    If IsCdLin(Src(J)) Then WNewLno = J + 1: Exit Function
Next
Stop
End Function
Private Function WNewLno1&(Src, IFm&, ITo&)
Dim J&
For J = IFm To ITo
    If Not XHas_Sfx(Src(J), " _") Then
        WNewLno1 = J + 1
        Exit Function
    End If
Next
Stop
End Function
Private Function WNeedIns(IsUsingCSub As Boolean, NewCSub$, OldCSub$) As Boolean
If Not IsUsingCSub Then Exit Function
If NewCSub = OldCSub Then Exit Function
WNeedIns = True
End Function

Private Function WNeedDlt(IsUsingCSub As Boolean, NewCSub$, OldCSub$) As Boolean
If OldCSub = "" Then Exit Function
If IsUsingCSub Then
    WNeedDlt = NewCSub <> OldCSub
Else
    WNeedDlt = OldCSub <> ""
End If
End Function

Private Function WIsUsingCSub(Src, IFm&, ITo&) As Boolean
Dim J&
For J = IFm To ITo
    If HasSubStrAy(Src(J), WIsUsingCSub1) Then WIsUsingCSub = True
Next
End Function
Private Property Get WIsUsingCSub1() As String()
Static O$()
If Sz(O) = 0 Then
Const A$ = " CSub" & ","
Const B$ = "(CSub" & ","
Const C$ = ", CSub"
Const D$ = "NevXThw CSub"
O = Ap_Sy(A, B, C, D)
End If
WIsUsingCSub1 = O
End Property

