Attribute VB_Name = "MVb_Er_Dmp"
Option Compare Binary
Option Explicit
Sub D(Optional A)
Select Case True
Case IsMissing(A): Debug.Print
Case IsArray(A): Ay_XDmp A
Case IsDic(A):   DicDmp CvDic(A)
Case IsRel(A):   RelDmp CvRel(A)
Case IsASet(A):   ASet_XDmp CvSet(A)
Case Else: Debug.Print A
End Select
End Sub
Sub Dmp(A)
D A
End Sub
Sub DmpTy(A)
Debug.Print TypeName(A)
End Sub

Sub Ay_XDmp(A, Optional WithIx As Boolean)
If Sz(A) = 0 Then Exit Sub
Dim I
If WithIx Then
    Dim J&
    For Each I In A
        Debug.Print J; ": "; I
        J = J + 1
    Next
Else
    For Each I In A
        Debug.Print I
    Next
End If
End Sub

Sub ChkEq(A, B)
If Not IsEq(A, B) Then
    Debug.Print "["; A; "] [" & TypeName(A) & "]"
    Debug.Print "["; B; "] [" & TypeName(A) & "]"
    Stop
End If
End Sub
