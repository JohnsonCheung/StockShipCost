Attribute VB_Name = "MVb_Dic"
Option Compare Binary
Option Explicit

Function CvDic(A) As Dictionary
Set CvDic = A
End Function

Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function

Function DicAddAy(A As Dictionary, Dy() As Dictionary) As Dictionary
Set DicAddAy = DicClone(A)
Dim J%
For J = 0 To UB(Dy)
   PushDic DicAddAy, Dy(J)
Next
End Function

Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set DicAddKeyPfx = O
End Function

Sub Dic_XIup(A As Dictionary, K$, V, Sep$) 'Iup means Ins_or_Upd
If A.Exists(K) Then
    A(K) = A(K) & Sep & V
Else
    A.Add K, V
End If
End Sub

Function DicAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicAllKeyIsNm = True
End Function

Function DicAllKeyIsStr(A As Dictionary) As Boolean
DicAllKeyIsStr = Ay_IsAllStr(A.Keys)
End Function

Function DicAllValIsStr(A As Dictionary) As Boolean
DicAllValIsStr = Ay_IsAllStr(A.Items)
End Function

Function DicAyKy(A() As Dictionary) As Variant()
Dim I
For Each I In AyNz(A)
   PushNoDupAy DicAyKy, CvDic(I).Keys
Next
End Function

Function DicByDry(DicDry) As Dictionary
Dim O As New Dictionary
If Sz(DicDry) <> 0 Then
   Dim Dr
   For Each Dr In DicDry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DicByDry = O
End Function

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set DicClone = O
End Function

Function DicDr(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DicDr = O
End Function

Function DicDRs_Fny(InclDicValTy As Boolean) As String()
DicDRs_Fny = SplitSpc("Key Val"): If InclDicValTy Then PushI DicDRs_Fny, "ValTy"
End Function

Function DicDry(A As Dictionary, Optional InclDicValTy As Boolean) As Variant()
Dim I, Dr
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Sz(K) = 0 Then Exit Function
For Each I In K
    If InclDicValTy Then
        Dr = Array(I, A(I), TypeName(A(I)))
    Else
        Dr = Array(I, A(I))
    End If
    Push DicDry, Dr
Next
End Function

Function DicIntersect(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicIntersect = O
End Function

Function DicIsEmp(A As Dictionary) As Boolean
DicIsEmp = A.Count = 0
End Function

Sub Dic_IsEq_XAss(A As Dictionary, B As Dictionary, Fun$, Optional N1$ = "A", Optional N2$ = "B")
If Not Dic_IsEq(A, B) Then XThw Fun, "2 given dic are diff", QQ_Fmt("[?] [?]", N1, N2), Dic_Fmt(A), Dic_Fmt(B)
End Sub

Function DicIsLinesDic(A As Dictionary) As Boolean
If Not DicIsStrDic(A) Then Exit Function
If Not Ay_IsLinesAy(A.Items) Then Exit Function
End Function

Function DicIsStrDic(A As Dictionary) As Boolean
If Not DicHasStrKey(A) Then Exit Function
DicIsStrDic = Ay_IsAllStr(A.Items)
End Function

Function DicKVLy(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, W%, Ky
Ky = A.Keys
W = Ay_Wdt(Ky)
For Each K In Ky
   Push O, XAlignL(K, W) & " = " & A(K)
Next
DicKVLy = O
End Function

Function DicKeySy(A As Dictionary) As String()
DicKeySy = Ay_Sy(A.Keys)
End Function

Function DicKyJnVal$(A As Dictionary, Ky, Optional Sep$ = vbCrLf & vbCrLf)
Dim O$(), K
For Each K In AyNz(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
DicKyJnVal = Join(O, Sep)
End Function

Function DicKySy(A As Dictionary, Ky$()) As String()
Dim K
For Each K In AyNz(Ky)
    PushI DicKySy, A(K)
Next
End Function

Function DicLblLy(A As Dictionary, Lbl$) As String()
PushI DicLblLy, Lbl
PushI DicLblLy, vbTab & "Count=" & A.Count
PushIAy DicLblLy, Ay_XAdd_Pfx(Dic_Fmt(A, InclValTy:=True), vbTab)
End Function

Function DicLines(A As Dictionary) As String
DicLines = JnCrLf(Dic_Fmt(A))
End Function

Function Dic_Ly(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = Ay_XAlign_L(Key)
Dim J&
For J = 0 To UB(Key)
   O(J) = O(J) & " " & A(Key(J))
Next
Dic_Ly = O
End Function

Function DicLy2(A As Dictionary) As String()
Dim O$(), K
If A.Count = 0 Then Exit Function
For Each K In A.Keys
    Push O, DicLy2__1(CStr(K), A(K))
Next
DicLy2 = O
End Function

Function DicLy2__1(K$, Lines$) As String()
Dim O$(), J&
Dim Ly$()
    Ly = SplitCrLf(Lines)
For J = 0 To UB(Ly)
    Dim Lin$
        Lin = Ly(J)
        If XTak_FstChr(Lin) = " " Then Lin = "~" & XRmv_FstChr(Lin)
    Push O, K & " " & Lin
Next
DicLy2__1 = O
End Function

Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, Run(ValMapFun, A(K))
Next
Set DicMap = O
End Function

Function DicMaxValSz%(A As Dictionary)
'MthDic is DicOf_MthNm_zz_MthLinesAy
'MaxMthCnt is max-of-#-of-method per MthNm
Dim O%, K
For Each K In A.Keys
    O = Max(O, Sz(A(K)))
Next
DicMaxValSz = O
End Function

Function DicMge(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = Ssl_Sy(PfxSsl)
   Ny = Ay_XAdd_Sfx(Ny, "@")
If Sz(Av) <> Sz(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, DicAddKeyPfx(A, Ny(J))
   Next
Set DicMge = DicAddAy(A, Dy)
End Function

Function Dic_XMinus(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set Dic_XMinus = New Dictionary: Exit Function
If B.Count = 0 Then Set Dic_XMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set Dic_XMinus = O
End Function

Function DicSelIntoAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntoAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = Ay_Sy(DicSelIntoAy(A, Ky))
End Function

Function DicStrKy(A As Dictionary) As String()
DicStrKy = Ay_Sy(A.Keys)
End Function

Function DicSwapKV(A As Dictionary) As Dictionary
Dim K
Set DicSwapKV = New Dictionary
For Each K In A.Keys
    DicSwapKV.Add A(K), K
Next
End Function

Function DicTy(A As Dictionary) As Dictionary
Set DicTy = DicMap(A, "TyNm")
End Function

Sub DicTyBrw(A As Dictionary)
Dic_XBrw DicTy(A)
End Sub

Function DicVal(A As Dictionary, K)
If IsNothing(A) Then Exit Function
If A.Exists(K) Then Asg A(K), DicVal
End Function

Function LikssDicKey$(A As Dictionary, Str)
Dim Likss$, K
For Each K In A
    Likss = A(K)
    If IsInLikss(Str, Likss) Then LikssDicKey = K: Exit Function
Next
End Function

Function MapStrDic(A) As Dictionary
Set MapStrDic = S1S2AyStrDic(A)
End Function

Function MayDicVal(MayDic As Dictionary, K)
If Not IsNothing(MayDic) Then MayDicVal = DicVal(MayDic, K)
End Function

Private Sub Z_DicMaxValSz()
Dim D As Dictionary, M%
'Set D = PjMthDic(CurPj)
M = DicMaxValSz(D)
Stop
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As Dictionary
Dim C() As Dictionary
Dim D$
Dim E$()
Dim F As Boolean
Dim G()
CvDic A
CvDicAy A
DicAddAy B, C
DicAddKeyPfx B, A
Dic_XIup B, D, A, D
DicAllKeyIsNm B
DicAllKeyIsStr B
DicAllValIsStr B
DicAyKy C
DicByDry A
DicClone B
DicDr B, E
DicDRs_Fny F
DicDry B, F
DicIntersect B, B
DicIsEmp B
Dic_IsEq_XAss B, B, D, D, D
DicIsLinesDic B
DicIsStrDic B
DicKVLy B
DicKeySy B
DicKyJnVal B, A, D
DicKySy B, E
DicLblLy B, D
DicLines B
DicLy2 B
DicLy2__1 D, D
DicMap B, D
DicMaxValSz B
DicMge B, D, G
Dic_XMinus B, B
DicSelIntoAy B, E
DicSelIntoSy B, E
DicStrKy B
DicSwapKV B
DicTy B
DicTyBrw B
DicVal B, A
LikssDicKey B, A
MapStrDic A
MayDicVal B, A
End Sub

Private Sub Z()
Z_DicMaxValSz
End Sub
