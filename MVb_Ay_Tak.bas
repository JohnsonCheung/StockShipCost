Attribute VB_Name = "MVb_Ay_Tak"
Option Compare Binary
Option Explicit

Function Ay_XTak_BefDD(A) As String()
Ay_XTak_BefDD = AyMap_Sy(A, "XTak_BefDD")
End Function

Function Ay_XTak_AftDot(A) As String()
Dim I
For Each I In AyNz(A)
    Push Ay_XTak_AftDot, XTak_AftDot(A)
Next
End Function

Function Ay_XTak_Aft(A, Sep$) As String()
Dim I
For Each I In AyNz(A)
    PushI Ay_XTak_Aft, XTak_Aft(I, Sep)
Next
End Function

Function Ay_XTak_Bef(A, Sep$) As String()
Dim I
For Each I In AyNz(A)
    PushI Ay_XTak_Bef, XTak_Bef(I, Sep)
Next
End Function

Function Ay_XTak_BefDot(A) As String()
Dim X
For Each X In AyNz(A)
    PushI Ay_XTak_BefDot, XTak_BefDot(X)
Next
End Function

Function Ay_XTak_BefOrAll(A, Sep$) As String()
Dim I
For Each I In AyNz(A)
    Push Ay_XTak_BefOrAll, XTak_BefOrAll(I, Sep)
Next
End Function

Function Ay_XTak_T1(A) As String()
Dim L
For Each L In AyNz(A)
    PushI Ay_XTak_T1, Lin_T1(L)
Next
End Function

Function Ay_XTak_T2(A) As String()
Ay_XTak_T2 = AyMap_Sy(A, "Lin_T2")
End Function

Function Ay_XTak_T3(A) As String()
Ay_XTak_T3 = AyMap_Sy(A, "Lin_T3")
End Function
