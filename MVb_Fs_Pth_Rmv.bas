Attribute VB_Name = "MVb_Fs_Pth_Rmv"
Option Compare Binary
Option Explicit
Private Sub Z_Pth_XRmv_EmpSubDirR()
Debug.Print "Before-----"
D Pth_EmpPthAyR(TmpRoot)
Pth_XRmv_EmpSubDirR TmpRoot
Debug.Print "After-----"
D Pth_EmpPthAyR(TmpRoot)
End Sub
Sub Pth_XRmv_EmpSubDirR(A)
Dim Ay$(), I, J%
Lp:
    J = J + 1: If J > 10000 Then Stop
    Ay = Pth_EmpPthAyR(A): If Sz(Ay) = 0 Then Exit Sub
    For Each I In Ay
        RmDir I
    Next
    GoTo Lp
End Sub

Sub Pth_XRmv_EmpSubDir(A$)
Dim I
For Each I In AyNz(Pth_PthAy(A))
   Pth_XRmv_IfEmp CStr(I)
Next
End Sub

Sub Pth_XRmv_IfEmp(A$)
If Not PthIsExist(A) Then Exit Sub
If PthIsEmp(A) Then Exit Sub
RmDir A
End Sub



