Attribute VB_Name = "MVb_Ay_X_Pred"
Option Compare Binary
Option Explicit

Function AyPred_AllTrue(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_AllTrue = ItrPred_IsAllTrue(A, Pred)
End Function

Function AyPred_SomFalse(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_SomFalse = ItrPred_IsSomFalse(A, Pred)
End Function

Sub AyPred_SplitAsg(A, Pred$, OTrueAy, OFalseAy)
Dim O1, O2
O1 = Ay_XCln(A)
O2 = O1
Dim X
For Each X In AyNz(A)
    If Run(Pred, X) Then
        Push OTrueAy, X
    Else
        Push OFalseAy, X
    End If
Next
End Sub

Function AyPred_XHasSomTrue(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_XHasSomTrue = Itr_XHas_PredTrue(A, Pred)
End Function

Function AyPred_IsAllFalse(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_IsAllFalse = ItrPred_IsAllFalse(A, Pred)
End Function
