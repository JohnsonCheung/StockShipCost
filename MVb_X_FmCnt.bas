Attribute VB_Name = "MVb_X_FmCnt"
Option Compare Binary
Option Explicit

Function CvFmCnt(A) As FmCnt
Set CvFmCnt = A
End Function

Function New_FmCnt(FmLno, Cnt) As FmCnt
Dim O As New FmCnt
Set New_FmCnt = O.Init(FmLno, Cnt)
End Function

Function FmCntAy_IsEq(A() As FmCnt, B() As FmCnt) As Boolean
If Sz(A) <> Sz(B) Then Exit Function
Dim X, J&
For Each X In AyNz(A)
    If Not FmCntIsEq(CvFmCnt(X), B(J)) Then Exit Function
    J = J + 1
Next
FmCntAy_IsEq = True
End Function

Function FmCntAy_IsInOrd(A() As FmCnt) As Boolean
Dim J%
For J = 0 To UB(A) - 1
    With A(J)
        If .FmLno = 0 Then Exit Function
        If .Cnt = 0 Then Exit Function
        If .FmLno + .Cnt > A(J + 1).FmLno Then Exit Function
    End With
Next
FmCntAy_IsInOrd = True
End Function

Function FmCntAyLinCnt%(A() As FmCnt)
Dim I, C%, O%
For Each I In A
    C = CvFmCnt(I).Cnt
    If C > 0 Then O = O + C
Next
FmCntAyLinCnt = O
End Function

Function FmCntAyLy(A() As FmCnt) As String()
Dim I
For Each I In AyNz(A)
    PushI FmCntAyLy, FmCntStr(CvFmCnt(I))
Next
End Function

Function FmCntIsEq(A As FmCnt, B As FmCnt) As Boolean
With A
    If .FmLno <> B.FmLno Then Exit Function
    If .Cnt <> B.Cnt Then Exit Function
End With
FmCntIsEq = True
End Function

Function FmCntStr$(A As FmCnt)
FmCntStr = "FmLno[" & A.FmLno & "] Cnt[" & A.Cnt & "]"
End Function

Private Sub ZZ()
Dim A As Variant
Dim B() As FmCnt
Dim C As FmCnt
CvFmCnt A
New_FmCnt A, A
FmCntAy_IsEq B, B
FmCntAy_IsInOrd B
FmCntAyLinCnt B
FmCntAyLy B
FmCntIsEq C, C
FmCntStr C
End Sub

Private Sub Z()
End Sub
