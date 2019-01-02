Attribute VB_Name = "MVb_Tst_Dic"
Option Explicit
Private Sub Can_a_Dic_with_ay_val_be_PUSHED()
Dim A As Dictionary, Act, V
GoSub T1
Exit Sub
T1: 'This fail
    Set A = New Dictionary
    A.Add "A", EmpAy
    PushI A("A"), 1
    V = A("A")
    Act = Sz(V)
    If Sz(Act) <> 1 Then Stop
    Return
T2:  'Should Pass
    Set A = New Dictionary
    A.Add "A", EmpAy
    V = A("A")
    PushI V, 1
    A("A") = V
    Act = A("A")
    If Sz(Act) <> 1 Then Stop
    Return
'Ans is: Cannot
End Sub
