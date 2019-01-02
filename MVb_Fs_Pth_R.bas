Attribute VB_Name = "MVb_Fs_Pth_R"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Fs_Pth_R."
Private O$(), A_Spec$, A_Atr As FileAttribute ' Used in Pth_PthAyR/Pth_FfnAyR

Function Pth_EmpPthAyR(A) As String()
Dim I
For Each I In AyNz(Pth_PthAyR(A))
    If PthIsEmp(I) Then PushI Pth_EmpPthAyR, I
Next
End Function

Function Pth_EntAyR(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Const CSub$ = CMod & "Pth_EntAyR"
Erase O
A_Spec = FilSpec
A_Atr = Atr
Pth_EntAyR1 A
Pth_EntAyR = O
End Function

Private Sub Pth_EntAyR1(A)
Ass PthIsExist(A)
If Sz(O) Mod 1000 = 0 Then Debug.Print "Pth_PthAyR1: (Each 1000): " & A
PushI O, A
PushIAy O, Pth_FfnAy(A, A_Spec, A_Atr)
Dim I, P$()
P = Pth_PthAy(A, A_Spec, A_Atr)
For Each I In AyNz(P)
    Pth_EntAyR1 I
Next
End Sub

Function Pth_FfnAyR(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
Pth_FfnAyR1 A
Pth_FfnAyR = O
End Function

Private Sub Pth_FfnAyR1(A)
PushIAy O, Pth_FfnAy(A, A_Spec, A_Atr)
If Sz(O) Mod 1000 = 0 Then Debug.Print "Pth_PthAyR1: (Each 1000): " & A
Dim P$(): P = Pth_PthAy(A, A_Spec, A_Atr)
If Sz(P) = 0 Then Exit Sub
Dim I
For Each I In P
    Pth_FfnAyR1 I
Next
End Sub

Function Pth_PthAyR(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
Pth_PthAyR1 A
Pth_PthAyR = O
End Function

Private Sub Pth_PthAyR1(A)
Dim P$(): P = Pth_PthAy(A, A_Spec, A_Atr)
If Sz(P) = 0 Then Exit Sub
If Sz(O) Mod 1000 = 0 Then Debug.Print "Pth_PthAyR1: (Each 1000): " & A
PushIAy O, P
Dim I
For Each I In P
    Pth_PthAyR1 I
Next
End Sub

Private Sub ZZ_Pth_EntAyR()
Dim A$(): A = Pth_EntAyR("C:\users\user\documents\")
Debug.Print Sz(A)
Stop
Ay_XDmp A
End Sub

Private Sub Z_Pth_EmpPthAyR()
Brw Pth_EmpPthAyR(TmpRoot)
End Sub

Private Sub Z_Pth_EntAyR()
Brw Pth_EntAyR(TmpRoot)
End Sub

Private Sub Z_Pth_FfnAyR()
D Pth_FfnAyR("C:\Users\User\Documents\WindowsPowershell\")
End Sub

Private Sub Z_Pth_XRmv_EmpSubDirR()
Pth_XRmv_EmpSubDirR TmpRoot
End Sub

Private Sub Z()
Z_Pth_EntAyR
Z_Pth_FfnAyR
Exit Sub
'Pth_EmpPthAyR
'Pth_EntAyR
'Pth_FfnAyR
'Pth_PthAyR
End Sub
