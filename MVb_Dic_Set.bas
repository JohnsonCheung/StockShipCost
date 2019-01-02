Attribute VB_Name = "MVb_Dic_Set"
Option Compare Binary
Option Explicit

Function CvSet(A) As ASet
Set CvSet = A
End Function

Property Get EmpASet() As ASet
Set EmpASet = New ASet
Set EmpASet.ASet = New Dictionary
End Property

Function CvASet(A) As ASet
Set CvASet = A
End Function

Function IsASet(A) As Boolean
IsASet = TypeName(A) = "ASet"
End Function

Function New_ASet(Optional Itr, Optional NoBlankStr As Boolean) As ASet
Dim O As ASet
Set O = EmpASet
If Not IsMissing(Itr) Then
    ASet_XAdd_Itr O, Itr, NoBlankStr
End If
Set New_ASet = O
End Function
Function New_ASet_AP(ParamArray Ap()) As ASet

End Function
Function New_ASet_SSL(Ssl) As ASet
Set New_ASet_SSL = New_ASet(Ssl_Sy(Ssl))
End Function

Function ASet_XAdd(A As ASet, B As ASet) As ASet
Set ASet_XAdd = EmpASet
ASet_Push ASet_XAdd, A
ASet_Push ASet_XAdd, A
End Function

Sub ASet_XRmv_Itm(O As ASet, Itm)
If ASet_XHas(O, Itm) Then O.ASet.Remove Itm
End Sub
Sub ASet_XPush(O As ASet, Itm)
If Not ASet_XHas(O, Itm) Then O.ASet.Add Itm, Empty
End Sub

Sub ASet_XAdd_Itr(O As ASet, Itr, Optional NoBlankStr As Boolean)
Dim I
If NoBlankStr Then
    For Each I In AyNz(Itr)
        If I <> "" Then
            ASet_XPush O, I
        End If
    Next
Else
    For Each I In Itr
        ASet_XPush O, I
    Next
End If
End Sub

Function ASet_Clone(A As ASet) As ASet
Dim O As ASet
O = EmpASet
ASet_XAdd_Itr O, A.ASet.Keys
ASet_Clone = O
End Function

Function ASet_Cnt&(A As ASet)
ASet_Cnt = A.ASet.Count
End Function

Sub ASet_XDmp(A As ASet)
D A.ASet.Keys
End Sub

Sub ASet_XBrw(A As ASet, Optional Fnn$)
Brw A.ASet.Keys, Fnn
End Sub

Function ASet_Fmt(A As ASet) As String()
ASet_Fmt = Ap_Sy(A.ASet.Keys)
End Function

Function ASet_XMinus(A As ASet, B As ASet) As ASet
Dim O As ASet, I
Set O = EmpASet
For Each I In ASet_Itms(A)
    If Not ASet_XHas(B, I) Then ASet_XPush O, I
Next
Set ASet_XMinus = O
End Function
Function ASet_XHas(A As ASet, Itm) As Boolean
ASet_XHas = A.ASet.Exists(Itm)
End Function

Function ASet_IsEq(A As ASet, B As ASet) As Boolean
If ASet_Cnt(A) <> ASet_Cnt(B) Then Exit Function
Dim K
For Each K In ASet_Itms(A)
    If Not ASet_XHas(B, K) Then Exit Function
Next
ASet_IsEq = True
End Function

Sub ASet_IsEq_XAss(A As ASet, B As ASet, Optional Msg$ = "Two set are diff", Optional ANm$ = "Set-A", Optional BNm$ = "Set-B")
If ASet_IsEq(A, B) Then Exit Sub
Dim O$()
PushI O, Msg
PushI O, QQ_Fmt("?-Cnt(?) / ?-Cnt(?)", ANm, ASet_Cnt(A), BNm, ASet_Cnt(B))
PushI O, ANm & " --------------------"
PushIAy O, ASet_Fmt(A)
PushI O, BNm & " --------------------"
PushIAy O, ASet_Fmt(B)
Ay_XBrw_XHalt O
End Sub

Function ASet_IsEmp(A As ASet) As Boolean
ASet_IsEmp = A.ASet.Count = 0
End Function

Function ASet_IsEq_IN_ORD(A As ASet, B As ASet) As Boolean
ASet_IsEq_IN_ORD = Ay_IsEq(A.ASet.Keys, B.ASet.Keys)
End Function
Function ASet_FstItm(A As ASet)
If ASet_IsEmp(A) Then XThw_Msg CSub, "Given ASet is empty"
Dim I
For Each I In ASet_Itms(A)
    Asg I, ASet_FstItm
Next
End Function
Function ASet_AbcDic(A As ASet) As Dictionary
'AbcDic means the keys is comming from ASet the value is starting from A, B, C
Dim O As New Dictionary, J%, K
For Each K In ASet_Itms(A)
    O.Add K, Chr(65 + J%)
    J = J + 1
Next
Set ASet_AbcDic = O
End Function
Function ASet_Itms(A As ASet)
ASet_Itms = A.ASet.Keys
End Function

Function ASet_Lin$(A As ASet)
ASet_Lin = JnSpc(A.ASet.Keys)
End Function

Sub ASet_Push(O As ASet, A As ASet)
ASet_XAdd_Itr O, ASet_Itms(A)
End Sub

Function ASet_Sy(A As ASet) As String()
ASet_Sy = Ay_Sy(A.ASet.Keys)
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As ASet
Dim C$
CvSet A
IsASet A
New_ASet A
New_ASet_SSL A
ASet_XAdd B, B
ASet_XPush B, A
ASet_XAdd_Itr B, A
ASet_Clone B
ASet_Cnt B
ASet_XDmp B
ASet_Fmt B
ASet_XHas B, A
ASet_IsEq B, B
ASet_IsEq_XAss B, B, C, C, C
ASet_IsEq_IN_ORD B, B
ASet_Itms B
End Sub

Private Sub Z()
End Sub
