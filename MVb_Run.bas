Attribute VB_Name = "MVb_Run"
Option Compare Binary
Option Explicit
Function Pipe(Pm, MthNy0)
Dim O: Asg Pm, O
Dim I
For Each I In CvNy(MthNy0)
   Asg Run(I, O), O
Next
Asg O, Pipe
End Function

Function RunAv(MthNm$, Av)
Dim O
Select Case Sz(Av)
Case 0: O = Run(MthNm)
Case 1: O = Run(MthNm, Av(0))
Case 2: O = Run(MthNm, Av(0), Av(1))
Case 3: O = Run(MthNm, Av(0), Av(1), Av(2))
Case 4: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Stop
End Select
RunAv = O
End Function

Function Fcmd_XRun_Max$(A$, ParamArray PmAp())
' WinSty As VbAppWinStyle = vbMaximizedFocus)
Dim Av(): Av = PmAp
Dim Lin$
    Lin = JnSpc(AyQuoteDbl(Ay_XAdd_(Array(A), Av)))
Shell Lin, vbMaximizedFocus
Fcmd_XRun_Max = A
End Function

Private Sub ZZ_Fcmd_Run_Max()
Fcmd_XRun_Max "Cmd"
MsgBox "AA"
End Sub
