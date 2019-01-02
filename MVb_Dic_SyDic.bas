Attribute VB_Name = "MVb_Dic_SyDic"
Option Compare Binary
Option Explicit

Function SyDicClone(SyDic As Dictionary) As Dictionary
IsSyDicAss SyDic, CSub
Dim Sy$(), K
Set SyDicClone = New Dictionary
For Each K In SyDic.Keys
    Sy = SyDic(K)
    SyDicClone.Add K, Sy
Next
End Function
Function IsSyDic(A) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
Set D = A
For Each I In D.Keys
    V = D(I)
    If Not IsSy(V) Then Exit Function
Next
IsSyDic = True
End Function

Sub IsSyDicAss(A As Dictionary, Fun$)
If Not IsSyDic(A) Then XThw Fun, "Given dictionary is not SyDic, all key is string and val is Sy", "Give-Dictionary", Dic_Fmt(A)
End Sub
