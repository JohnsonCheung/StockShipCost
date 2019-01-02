Attribute VB_Name = "MVb__Obj"
Option Compare Binary
Option Explicit
Function Obj_Prp_PTH$(Obj, PrpSsl$)
Dim Ny$(): Ny = Ssl_Sy(PrpSsl)
Dim O$(), I
For Each I In Ny
    Push O, CallByName(Obj, CStr(I), VbGet)
Next
Obj_Prp_PTH = Join(O, "|")
End Function

Function ObjDr(A, PrpNy0) As Variant()
Dim PrpNy$(), U%, O(), J%
PrpNy = CvNy(PrpNy0)
U = UB(PrpNy)
ReDim O(U)
For J = 0 To U
    Asg Obj_Prp(A, PrpNy(J)), O(J)
Next
ObjDr = O
End Function

Function Any_Obj_NmPfx(O, NmPfx$) As Boolean
Any_Obj_NmPfx = XHas_Pfx(Obj_Nm(O), NmPfx)
End Function
Function Obj_IsEq(A, B) As Boolean
Obj_IsEq = ObjPtr(A) = ObjPtr(B)
End Function


Function Obj_Nm$(A)
If IsNothing(A) Then Obj_Nm = "#nothing#": Exit Function
On Error GoTo X
Obj_Nm = A.Name
Exit Function
X:
Obj_Nm = "#" & Err.Description & "#"
End Function

Function Obj_Prp(A, P)
If IsNothing(A) Then XDmp_Ly CSub, "Given object is nothing", "PrpNm", P: Exit Function
On Error GoTo X
Asg CallByName(A, P, VbGet), Obj_Prp
Exit Function
X:
Dim E$
E = Err.Description
XDmp_Ly CSub, "Error in getting Obj-Prp", "Obj-TypeName PrpNm Er", TypeName(A), P, E
End Function

Function Obj_PrpAy(A, PrpNy0) As Variant()
If IsNothing(A) Then XDmp_Ly CSub, "Given object is nothing", "PrpNy0", PrpNy0: Exit Function
Dim I
For Each I In CvNy(PrpNy0)
    Push Obj_PrpAy, Obj_Prp(A, I)
Next
End Function

Function Obj_PrpDr(A, PrpNy0) As Variant()
Obj_PrpDr = Obj_PrpAy(A, PrpNy0)
End Function

Function Obj_PrpPth(A, PrpPth$)
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim P$()
    P = Split(PrpPth, ".")
Dim O
    Dim J%, U%
    Set O = A
    U = UB(P)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, P(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next

Asg CallByName(O, P(U), VbGet), Obj_PrpPth ' Last Prp may be non-object, so must use 'Asg'
End Function

Function Obj_ToStr$(A)
On Error GoTo X
Obj_ToStr = A.ToStr: Exit Function
X: Obj_ToStr = XQuote_SqBkt(TypeName(A))
End Function

Private Sub ZZZ_Obj_Prp_PTH()
Dim Act$: Act = Obj_Prp_PTH(excel.Application.Vbe.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub
