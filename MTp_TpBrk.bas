Attribute VB_Name = "MTp_TpBrk"
Option Compare Binary
Option Explicit
Public Type TpSec
    Nm As String
    GpAy() As Gp
End Type
Public Type TpBrk
    Er() As String
    RmkDic As New Dictionary
    SecAy() As TpSec
End Type
Function TpT1LyDic(A) As Dictionary

End Function

Function Tp_TpBrk(A$) As TpBrk
', OErLy$(), ORmkDic As Dictionary, Ny0, ParamArray OLyAp())
Dim O(), J%, U%
'O = ClnBrk1(LyCln(SplitCrLf(A)), Ny0)
U = UB(O)
For J = 0 To U - 2
    'OLyAp(J) = O(J)
Next
'OErLy = O(U + 1)
'Set ORmkDic = O(U + 2)
End Function

Function Ly_LnxAy(A$()) As Lnx()
Dim J&, O() As Lnx
If Sz(A) = 0 Then Exit Function
For J = 0 To UB(A)
    PushObj O, Lnx(J, A(J))
Next
Ly_LnxAy = O
End Function
