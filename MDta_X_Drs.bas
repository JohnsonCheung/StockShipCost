Attribute VB_Name = "MDta_X_Drs"
Option Compare Binary
Option Explicit

Function CvDrs(A) As Drs
Set CvDrs = A
End Function

Function New_Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set New_Drs = O.Init(CvNy(Fny0), Dry)
End Function

Function Drs_XAdd_Col(A As Drs, ColNm$, ColVal) As Drs
Dim Fny$(): Fny = A.Fny
Dim NewFny$(): NewFny = Fny: PushI NewFny, ColNm
Set Drs_XAdd_Col = New_Drs(NewFny, Dry_XAdd_Col(A.Dry, ColVal))
End Function

Function Drs_XAdd_ConstCol(A As Drs, ColNm$, ConstVal) As Drs
Dim Fny$()
    Fny = A.Fny
    Push Fny, ColNm
Set Drs_XAdd_ConstCol = New_Drs(Fny, Dry_XAdd_ConstCol(A.Dry, ConstVal))
End Function

Function Drs_XAdd_RowIxCol(A As Drs) As Drs
Dim Fny$()
Dim Dry()
    Fny = AyIns(A.Fny, "Ix")
    Dim J&, Dr
    For Each Dr In AyNz(A.Dry)
        Dr = AyIns(Dr, J): J = J + 1
        Push Dry, Dr
    Next
Set Drs_XAdd_RowIxCol = New_Drs(Fny, Dry)
End Function

Function Drs_XAdd_ValIdCol(A As Drs, ColNm$, Optional ColNmPfx$) As Drs
Dim Ix%, Fny$()
Fny = A.Fny
Ix = Ay_Ix(Fny, ColNm): If Ix = -1 Then Stop
    Dim X$, Y$, C$
        C = ColNmPfx & ColNm
        X = C & "Id"
        Y = C & "Cnt"
    If Ay_XHas(Fny, X) Then Stop
    If Ay_XHas(Fny, Y) Then Stop
    PushIAy Fny, Array(X, Y)
Set Drs_XAdd_ValIdCol = New_Drs(Fny, Dry_XAdd_ValIdCol(A.Dry, Ix))
End Function
Function IsDrs(A) As Boolean
IsDrs = TypeName(A) = "Drs"
End Function
Sub Drs_XBrw(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional Fnn$)
Ay_XBrw Drs_Fmt(A, MaxColWdt, BrkColNm$), Fnn
End Sub

Function DrsCol(A As Drs, ColNm) As Variant()
'DrsCol = DryColInto(A.Dry, ColNm)
End Function

Function DrsColInto(A As Drs, F, OInto)
Dim O, Ix%, Dry(), Dr
Ix = Ay_Ix(A.Fny, F): If Ix = -1 Then Stop
O = OInto
Erase O
Dry = A.Dry
If Sz(Dry) = 0 Then DrsColInto = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
DrsColInto = O
End Function

Function DrsColSy(A As Drs, F) As String()
DrsColSy = DrsColInto(A, F, EmpSy)
End Function

Sub Drs_XDmp(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$)
Ay_XDmp Drs_Fmt(A, MaxColWdt, BrkColNm$)
End Sub

Function Drs_XDrp_Col(A As Drs, FF) As Drs
Dim Fny$(): Fny = FF_Fny(FF)
XAss Ay_XHasSubAy(A.Fny, Fny), CSub, "Given FF has some field not in Drs.Fny", "FF Drs.Fny", FF, A.Fny
Dim IxAy&()
    IxAy = Ay_IxAy(A.Fny, Fny)
Dim Dry()
    Fny = Ay_XExl_IxAy(A.Fny, IxAy)
    Dry = Dry_XRmv_ColByIxAy(A.Dry, IxAy)
Set Drs_XDrp_Col = New_Drs(Fny, Dry)
End Function

Function Drs_Dt(A As Drs, DtNm$) As Dt
Set Drs_Dt = New_Dt(DtNm, A.Fny, A.Dry)
End Function

Function Drs_FF$(A As Drs)
Drs_FF = Fny_FF(A.Fny)
End Function

Function DrsInsCol(A As Drs, ColNm$, C) As Drs
Set DrsInsCol = New_Drs(AyIns(A.Fny, ColNm), Dry_XIns_Col(A.Dry, C))
End Function

Function DrsInsColAft(A As Drs, C$, FldNm$) As Drs
Set DrsInsColAft = DrsInsColXxx(A, C, FldNm, True)
End Function

Function DrsInsColBef(A As Drs, C$, FldNm$) As Drs
Set DrsInsColBef = DrsInsColXxx(A, C, FldNm, False)
End Function

Private Function DrsInsColXxx(A As Drs, C$, FldNm$, IsAft As Boolean) As Drs
Dim Fny$(), Dry(), Ix&, Fny1$()
Fny = A.Fny
Ix = Ay_Ix(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = AyIns(Fny, FldNm, CLng(Ix))
Dry = Dry_XIns_Col(A.Dry, Ix)
Set DrsInsColXxx = New_Drs(Fny1, Dry)
End Function

Function DrsIsEq(A As Drs, B As Drs) As Boolean
If Not Ay_IsEq(A.Fny, B.Fny) Then Exit Function
If Not DryIsEq(A.Dry, B.Dry) Then Exit Function
DrsIsEq = True
End Function

Function DrsKeyCntDic(A As Drs, K$) As Dictionary
Dim Dry(), O As New Dictionary, Fny$(), Dr, Ix%, KK$
Fny = A.Fny
Ix = Ay_Ix(Fny, K)
Dry = A.Dry
If Sz(Dry) > 0 Then
    For Each Dr In A.Dry
        KK = Dr(Ix)
        If O.Exists(KK) Then
            O(KK) = O(KK) + 1
        Else
            O.Add KK, 1
        End If
    Next
End If
Set DrsKeyCntDic = O
End Function

Function New_Drs_ByDRs_Lines(A$) As Drs
Set New_Drs_ByDRs_Lines = New_Drs_ByDRs_Ly(SplitLines(A))
End Function

Function New_Drs_ByDRs_Ly(A$()) As Drs
Dim L, Dry()
For Each L In AyNz(A)
    PushI Dry, Lin_TermAy(L)
Next
Set New_Drs_ByDRs_Ly = New_Drs(Lin_TermAy(A(0)), Dry)
End Function

Function Drs_NCol%(A As Drs)
Drs_NCol = Max(Sz(A.Fny), Dry_NCol(A.Dry))
End Function

Function DrsNRow&(A As Drs)
DrsNRow = Sz(A.Dry)
End Function

Function DrsPkDiff(A As Drs, B As Drs, PkSs$) As Drs

End Function

Function DrsPkMinus(A As Drs, B As Drs, PkSs$) As Drs
Dim Fny$(), PkIxAy&()
Fny = A.Fny: If Not Ay_IsEq(Fny, B.Fny) Then Stop
PkIxAy = Ay_IxAy(Fny, Ssl_Sy(PkSs))
Set DrsPkMinus = New_Drs(Fny, DryPkMinus(A.Dry, B.Dry, PkIxAy))
End Function

Function DrsReOrd(A As Drs, Partial_Fny0) As Drs
Dim ReOrdFny$(): ReOrdFny = CvNy(Partial_Fny0)
Dim IxAy&(): IxAy = Ay_IxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
Set DrsReOrd = New_Drs(OFny, ODry)
End Function

Function Drs_NRow_WH&(A As Drs, ColNm$, EqVal)
Drs_NRow_WH = DryRowCnt(A.Dry, Ay_Ix(A.Fny, ColNm), EqVal)
End Function

Function Drs_Sq(A As Drs) As Variant()
Dim NC&, NR&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NC = Max(Dry_NCol(Dry), Sz(Fny))
    NR = Sz(Dry)
Dim O()
ReDim O(1 To 1 + NR, 1 To NC)
Dim C&, R&, Dr
    For C = 1 To Sz(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NR
        Dr = Dry(R - 1)
        For C = 1 To Min(Sz(Dr), NC)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
Drs_Sq = O
End Function

Function DRs_StrCol_FstCol(A As Drs, ColNm) As String()
DRs_StrCol_FstCol = Ay_Sy(DrsCol(A, ColNm))
End Function

Function DrsSy(A As Drs, ColNm) As String()
DrsSy = DRs_StrCol_FstCol(A, ColNm)
End Function

Function DrsVblDrs(DrsVbl$) As Drs
Set DrsVblDrs = New_Drs_ByDRs_Ly(SplitVBar(DrsVbl))
End Function

Function DrsVbl_Drs(DrsVbl$) As Drs
'SpecStr:Vbl:VbarLine
'SpecStr:DrsVbl:Data-record-set-vbar-line
'DrsVbl_Drs = DRs_Ly_Drs(SplitVBar(DrsVbl))
Stop '
End Function

Function ItoDrs(A, PrpNy0) As Drs
Dim Fny$(): Fny = CvNy(PrpNy0)
Set ItoDrs = New_Drs(Fny, ItoSel(A, Fny))
End Function

Function LblSeqAy(A, N%) As String()
Dim O$(), J%, F$, L%
L = Len(N)
F = StrDup("0", L)
ReDim O(N - 1)
For J = 1 To N
    O(J - 1) = A & Format(J, F)
Next
LblSeqAy = O
End Function

Function LblSeqSsl$(A, N%)
LblSeqSsl = Join(LblSeqAy(A, N), " ")
End Function

Sub Drs_Push(O As Drs, A As Drs)
If IsNothing(O) Then
    Set O = A
    Exit Sub
End If
If IsNothing(A) Then Exit Sub
If Not IsEq(O.Fny, A.Fny) Then Stop
Set O = New_Drs(O.Fny, CvAy(AyAp_XAdd(O.Dry, A.Dry)))
End Sub

Private Sub ZZ_Drs_GpDic()
Dim Act As Dictionary, Dry(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dry = Array(Dr1, Dr2, Dr3)
Set Act = Dry_GpDic(Dry, Ap_IntAy(0), 2)
Ass Act.Count = 2
Ass Ay_IsEq(Act("A"), Array(1, 2))
Ass Ay_IsEq(Act("B"), Array(3))
Stop
End Sub

Private Sub ZZ_Drs_PivDrs()
Dim Act As Drs, Drs2 As Drs, Drs1 As Drs, N1%, N2%
'Set Drs1 = VbeFun12Drs(CurVbe)
N1 = Sz(Drs1.Dry)
'Set Drs2 = Vbe_Mth12Drs(CurVbe)
'N2 = Sz(Drs2.Dry)
'Debug.Print N1, N2
Set Act = Drs_PivDrs(Drs1, "Nm", "Lines")
Drs_XBrw Act
End Sub

Private Sub ZZ_Drs_PivDrs_1()
Dim Act As Drs, D As Drs, Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Set D = New_Drs("A B C", CvAy(Array(Dr1, Dr2, Dr3)))
Set Act = Drs_PivDrs(D, "A", "C")
Stop
Drs_XBrw Act
End Sub

Private Sub ZZ_DrsKeyCntDic()
Dim Drs As Drs, Dic As Dictionary
'Set Drs = Vbe_Mth12Drs(CurVbe)
Set Dic = DrsKeyCntDic(Drs, "Nm")
Dic_XBrw Dic
End Sub

Private Sub ZZ_Drs_XSel()
Drs_XBrw Drs_XSel(Samp_Drs1, "A B D")
End Sub

Private Property Get Z_Drs_Fmt()
GoTo ZZ
ZZ:
Ay_XDmp Drs_Fmt(Samp_Drs1)
End Property

Private Sub ZZ()
Dim A As Variant
Dim B()
Dim C As Drs
Dim D$
Dim E%
Dim F$()
CvDrs A
Drs_XAdd_Col C, D, A
Drs_XAdd_ConstCol C, D, A
Drs_XAdd_RowIxCol C
Drs_XAdd_ValIdCol C, D, D
Drs_XBrw C, E, D, D
DrsCol C, A
DrsColInto C, A, A
DrsColSy C, A
Drs_XDmp C, E, D
Drs_XDrp_Col C, A
Drs_Dt C, D
Drs_PivDrs C, D, D
DrsInsCol C, D, A
DrsInsColAft C, D, D
DrsInsColBef C, D, D
DrsIsEq C, C
DrsKeyCntDic C, D
DrsPkDiff C, C, D
DrsPkMinus C, C, D
DrsReOrd C, A
Drs_Sq C
DRs_StrCol_FstCol C, A
DrsSy C, A
DrsVblDrs D
DrsVbl_Drs D
ItoDrs A, A
LblSeqAy A, E
LblSeqSsl A, E
Drs_Push C, C
End Sub

Private Sub Z()
End Sub
