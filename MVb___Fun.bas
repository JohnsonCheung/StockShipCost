Attribute VB_Name = "MVb___Fun"
Option Compare Binary
Option Explicit
Public CSub$

Sub Asg(Fm, OTo)
If IsNumeric(OTo) Then
    If Not IsEmpty(OTo) Then
        OTo = Val(Fm)
        Exit Sub
    End If
End If
If IsObject(Fm) Then
    Set OTo = Fm
Else
    If IsNull(Fm) Then
        OTo = ""
    Else
        OTo = Fm
    End If
End If
End Sub

Sub Brw(A, Optional Fnn$)
Select Case True
Case IsStr(A): StrBrw A, Fnn
Case IsArray(A): Ay_XBrw A, Fnn
Case IsASet(A): ASet_XBrw CvASet(A), Fnn
Case IsDrs(A): Drs_XBrw CvDrs(A)
Case IsDic(A): Dic_XBrw CvDic(A)
Case Else: Stop
End Select
End Sub

Function CanCvLng(A) As Boolean
On Error GoTo X
Dim L&: L = CLng(A)
CanCvLng = True
X:
End Function

Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

Function InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Function
Next
InstrN = P
End Function

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) > O Then O = Av(J)
Next
Max = O
End Function

Function MaxVbTy(A As VbVarType, B As VbVarType) As VbVarType
If A = vbString Or B = vbString Then MaxVbTy = A: Exit Function
If A = vbEmpty Then MaxVbTy = B: Exit Function
If B = vbEmpty Then MaxVbTy = A: Exit Function
If A = B Then MaxVbTy = A: Exit Function
Dim AIsNum As Boolean, BIsNum As Boolean
AIsNum = IsVbTyNum(A)
BIsNum = IsVbTyNum(B)
Select Case True
Case A = vbBoolean And BIsNum: MaxVbTy = B
Case AIsNum And B = vbBoolean: MaxVbTy = A
Case A = vbDate Or B = vbDate: MaxVbTy = vbString
Case AIsNum And BIsNum:
    Select Case True
    Case A = vbByte: MaxVbTy = B
    Case B = vbByte: MaxVbTy = A
    Case A = vbInteger: MaxVbTy = B
    Case B = vbInteger: MaxVbTy = A
    Case A = vbLong: MaxVbTy = B
    Case B = vbLong: MaxVbTy = A
    Case A = vbSingle: MaxVbTy = B
    Case B = vbSingle: MaxVbTy = A
    Case A = vbDouble: MaxVbTy = B
    Case B = vbDouble: MaxVbTy = A
    Case A = vbCurrency Or B = vbCurrency: MaxVbTy = A
    Case Else: Stop
    End Select
Case Else: Stop
End Select
End Function

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = AyMin(Av)
End Function

Function NDig%(Length&)
Select Case True
Case 0 > Length: XThw CSub, "Length cannot <0", "Length", Length
Case 10 > Length: NDig = 1
Case 100 > Length: NDig = 2
Case 1000 > Length: NDig = 3
Case 10000 > Length: NDig = 4
Case 100000 > Length: NDig = 5
Case 1000000 > Length: NDig = 6
Case 10000000 > Length: NDig = 7
Case 100000000 > Length: NDig = 8
Case 1000000000 > Length: NDig = 9
Case Else: NDig = 10
End Select
End Function

Function New_Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
If Patn = "" Or Patn = "." Then Exit Function
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set New_Re = O
End Function

Sub Keys_XSnd(A$)
DoEvents
SendKeys A, True
End Sub

Sub XBrw(A, Optional Fnn$)
Brw A, Fnn
End Sub

Private Sub Z_InstrN()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InstrN(S, SubStr, N)
Ass Exp = Act
End Sub

Private Sub Z()
Z_InstrN
MVb__Fun:
End Sub
