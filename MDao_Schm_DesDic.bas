Attribute VB_Name = "MDao_Schm_DesDic"
Option Compare Binary
Option Explicit

Function DesDicFldDesDic(A As Dictionary, T, Fny$()) As Dictionary
Set DesDicFldDesDic = New Dictionary
Dim F, D$
For Each F In Fny
    D = ZDes(A, T, F)
    If D <> "" Then DesDicFldDesDic.Add F, D
Next
End Function

Private Function ZDes$(A As Dictionary, T, F)
'ZDes = Ap_JnDblDollar NOBLANK(
End Function
