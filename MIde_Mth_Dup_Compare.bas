Attribute VB_Name = "MIde_Mth_Dup_Compare"
Option Compare Binary
Option Explicit
Sub FunCmp(FunNm$, Optional InclSam As Boolean)
D FunCmpFmt(FunNm, InclSam)
End Sub

Function FunCmpFmt(FunNm, Optional InclSam As Boolean) As String()
'Found all Fun with given name and compare within curVbe if it is same
'Note: Fun is any-Mdy Fun/Sub/Prp-in-Md
Dim O$()
Dim N$(): ' N = FunFNmAy(FunNm)
DupMthFNy_ShwNotDupMsg N, FunNm
If Sz(N) <= 1 Then Exit Function
FunCmpFmt = Dic_CmpFmt(N, InclSam:=InclSam)
End Function

Private Sub Z_FunCmp()
FunCmp "Ffn_XDlt"
End Sub

Private Function Dic_CmpFmt(A, Optional OIx% = -1, Optional OSam% = -1, Optional InclSam As Boolean) As String()
'DupMthFNyGp is Variant/String()-of-MthFNm with all mth-nm is same
'MthFNm is MthNm in FNm-fmt
'          Mth is Prp/Sub/Fun in Md-or-Cls
'          FNm-fmt which is 'Nm:Pj.Md'
'DupMthFNm is 2 or more MthFNy with same MthNm
Ass DupMthFNyGp_IsVdt(A)
Dim J%, I%
Dim LinesAy$()
Dim UniqLinesAy$()
    LinesAy = AyMap_Sy(A, "FunFNm_MthLines")
    UniqLinesAy = Ay_XWh_Dist(LinesAy)
Dim MthNm$: MthNm = Brk(A(0), ":").S1
Dim Hdr$(): Hdr = Dic_CmpFmt__1Hdr(OIx, MthNm, Sz(A))
Dim Sam$(): Sam = Dic_CmpFmt__2Sam(InclSam, OSam, A, LinesAy)
Dim Syn$(): Syn = Dic_CmpFmt__3Syn(UniqLinesAy, LinesAy, A)
Dim Cmp$(): Cmp = Dic_CmpFmt__4Cmp(UniqLinesAy, LinesAy, A)
Dic_CmpFmt = AyAp_XAdd(Hdr, Sam, Syn, Cmp)
End Function

Private Function Dic_CmpFmt__1Hdr(OIx%, MthNm$, Cnt%) As String()
Dim O$(1)
O(0) = "================================================================"
Dim A$
    If OIx >= 0 Then A = QQ_Fmt("#DupMthNo(?) ", OIx): OIx = OIx + 1
O(1) = A + QQ_Fmt("DupMthNm(?) Cnt(?)", MthNm, Cnt)
Dic_CmpFmt__1Hdr = O
End Function

Private Function Dic_CmpFmt__2Sam(InclSam As Boolean, OSam%, DupMthFNyGp, LinesAy$()) As String()
If Not InclSam Then Exit Function
'{DupMthFNyGp} & {LinesAy} have same # of element
Dim O$()
Dim D$(): D = Ay_XWh_Dup(LinesAy)
Dim J%, X$()
For J = 0 To UB(D)
    X = Dic_CmpFmt__2Sam1(OSam, D(J), DupMthFNyGp, LinesAy)
    PushAy O, X
Next
Dic_CmpFmt__2Sam = O
End Function

Private Function Dic_CmpFmt__2Sam1(OSam%, Lines$, DupMthFNyGp, LinesAy$()) As String()
Dim A1$()
    If OSam > 0 Then
        Push A1, QQ_Fmt("#Sam(?) ", OSam)
        OSam = OSam + 1
    End If
Dim A2$()
    Dim J%
    For J = 0 To UB(LinesAy)
        If LinesAy(J) = Lines Then
            Push A2, "Shw """ & DupMthFNyGp(J) & """"
        End If
    Next
Dim A3$()
    A3 = Lines_BoxLy(Lines)
Dic_CmpFmt__2Sam1 = AyAp_XAdd(A1, A2, A3)
End Function

Private Function Dic_CmpFmt__3Syn(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim B$()
    Dim J%, I%
    Dim Lines
    For Each Lines In UniqLinesAy
        For I = 0 To UB(FunFNyGp)
            If Lines = LinesAy(I) Then
                Push B, FunFNyGp(I)
                Exit For
            End If
        Next
    Next
Dic_CmpFmt__3Syn = AyMapPX_Sy(B, "QQ_Fmt", "Sync_Fun ""?""")
End Function

Private Function Dic_CmpFmt__4Cmp(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim L2$() ' = From L1 with each element with MdDNm added in front
    ReDim L2(UB(UniqLinesAy))
    Dim Fnd As Boolean, DNm$, J%, Lines$, I%
    For J = 0 To UB(UniqLinesAy)
        Lines = UniqLinesAy(J)
        Fnd = False
        For I = 0 To UB(LinesAy)
            If LinesAy(I) = Lines Then
                DNm = FunFNyGp(I)
                L2(J) = DNm & vbCrLf & StrDup("-", Len(DNm)) & vbCrLf & Lines
                Fnd = False
                GoTo Nxt
            End If
        Next
        Stop
Nxt:
    Next
Dic_CmpFmt__4Cmp = LinesAy_LyPad(L2)
End Function

Function MthNmCmpFmt(A, Optional InclSam As Boolean) As String()
Dim N$(): N = MthNm_DupMthFNy(A)
If Sz(N) > 1 Then
    MthNmCmpFmt = Dic_CmpFmt(N, InclSam:=InclSam)
End If
End Function


Function VbeDupMthCmpLy(A As Vbe, B As WhPjMth, Optional InclSam As Boolean) As String()
Stop
Dim N$(): 'N = VbeDupMthFNm(A, B)
Dim Ay(): Ay = DupMthFNy_GpAy(N)
Dim O$(), J%
Push O, QQ_Fmt("Total ? dup function.  ? of them has mth-lines are same", Sz(Ay), DupMthFNyGpAy_AllSameCnt(Ay))
Dim Cnt%, Sam%
For J = 0 To UB(Ay)
    PushAy O, Dic_CmpFmt(Ay(J), Cnt, Sam, InclSam:=InclSam)
Next
VbeDupMthCmpLy = O
End Function



Private Sub ZZ_VbeDupMthCmpLy()
Brw VbeDupMthCmpLy(CurVbe, WhEmpPjMth)
End Sub




Private Sub Z()
Z_FunCmp
MIde_Mth_Dup_Compare:
End Sub
