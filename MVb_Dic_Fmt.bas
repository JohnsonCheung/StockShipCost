Attribute VB_Name = "MVb_Dic_Fmt"
Option Compare Binary
Option Explicit

Sub Dic_XBrw(A As Dictionary, Optional InclDicValTy As Boolean)
Ay_XBrw Dic_Fmt(A, InclDicValTy)
End Sub

Sub DicDmp(A As Dictionary, Optional InclDicValTy As Boolean, Optional Tit$ = "Key Val")
D Dic_Fmt(A, InclDicValTy, Tit)
End Sub
Function LyDicS1S2Ay(A As Dictionary) As S1S2()
Dim K
For Each K In A.Keys
    PushObj LyDicS1S2Ay, S1S2(K, JnCrLf(A(K)))
Next
End Function
Function Dic_Fmt(A As Dictionary, Optional InclValTy As Boolean, Optional Tit$ = "Key Val") As String()
If IsSyDic(A) Then
    Dic_Fmt = S1S2Ay_Fmt(LyDicS1S2Ay(A))
    Exit Function
End If
If ZHasLines(A) Or IsSyDic(A) Then
    Dic_Fmt = S1S2Ay_Fmt(DicS1S2Ay(A))
Else
    Dic_Fmt = ZLinFmt(A, InclValTy)
End If
End Function

Private Function ZHasLines(A As Dictionary) As Boolean
ZHasLines = True
Dim K
For Each K In A.Keys
    If IsLines(K) Then Exit Function
    If IsLines(A(K)) Then Exit Function
Next
ZHasLines = False
End Function

Private Function ZLinFmt(A As Dictionary, Optional InclDicValTy As Boolean) As String()
Dim K$(), O$(), V(), I, J&
If InclDicValTy Then ZLinFmt = ZLinFmt1(A): Exit Function
K = Ay_XAlign_L(A.Keys)
V = A.Items
For Each I In AyNz(K)
    PushI ZLinFmt, I & " " & Var_Lines(V(J))
    J = J + 1
Next
End Function

Private Function ZLinFmt1(A As Dictionary) As String()
Dim K, O$()
For Each K In A.Keys
    PushI O, K & " " & TypeName(A(K)) & " " & Var_Str(A(K))
Next
ZLinFmt1 = Ay_XAlign_2T(O)
End Function
