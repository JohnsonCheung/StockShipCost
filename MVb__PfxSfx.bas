Attribute VB_Name = "MVb__PfxSfx"
Option Compare Binary
Option Explicit

Function XAdd_Pfx$(A$, Pfx$)
XAdd_Pfx = Pfx & A
End Function

Function XAdd_PfxSfx$(A$, Pfx$, Sfx$)
XAdd_PfxSfx = Pfx & A & Sfx
End Function

Function XAdd_Sfx$(A$, Sfx$)
XAdd_Sfx = A & Sfx
End Function

Function XHas_PfxAy(A, PfxAy0, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
Dim I
For Each I In CvNy(PfxAy0)
   If XHas_Pfx(A, I, Cmp) Then XHas_PfxAy = True: Exit Function
Next
End Function
Function XAdd_PfxSpc_IfAny$(A)
If A = "" Then Exit Function
XAdd_PfxSpc_IfAny = " " & A
End Function

Function XHas_Pfx(A, Pfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
XHas_Pfx = StrComp(Left(A, Len(Pfx)), Pfx, Cmp) = 0
End Function

Function XHas_Spc(A) As Boolean
XHas_Spc = XHas_SubStr(A, " ")
End Function

Function XHas_SubStr(A, SubStr, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
XHas_SubStr = InStr(A, SubStr) > 0
End Function
Function XHas_PfxAp(A, ParamArray PfxAp()) As Boolean
Dim Av(): Av = PfxAp
XHas_PfxAp = XHas_PfxAy(A, Av)
End Function

Function XHas_PfxSpc(A, Pfx) As Boolean
XHas_PfxSpc = XHas_Pfx(A, Pfx & " ")
End Function
Function Ay_XAdd_Pfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), U&, J&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Pfx & A(J)
Next
Ay_XAdd_Pfx = O
End Function

Function Ay_XAdd_PfxSfx(A, Pfx, Sfx) As String()
Dim O$(), J&, U&
If Sz(A) = 0 Then Exit Function
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J) & Sfx
Next
Ay_XAdd_PfxSfx = O
End Function

Function Ay_XAdd_Sfx(A, Sfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = A(J) & Sfx
Next
Ay_XAdd_Sfx = O
End Function

Function Ay_IsAllEleXHas_Pfx(A, Pfx$) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
   If Not XHas_Pfx(I, Pfx) Then Exit Function
Next
Ay_IsAllEleXHas_Pfx = True
End Function


Function Ay_XAdd_CommaSpcSfxExlLas(A) As String()
Dim X, J, U%
U = UB(A)
For Each X In AyNz(A)
    If J = U Then
        Push Ay_XAdd_CommaSpcSfxExlLas, X
    Else
        Push Ay_XAdd_CommaSpcSfxExlLas, X & ", "
    End If
    J = J + 1
Next
End Function
Function XTak_SfxChr_InLis$(A, SfxChrLis$, Optional Cmp As VbCompareMethod = vbTextCompare)
If XHas_SfxChr(A, SfxChrLis, Cmp) Then XTak_SfxChr_InLis = XTak_LasChr(A)
End Function

Function XHas_SfxChr(A, SfxChrLis$, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
Dim J%
For J = 1 To Len(SfxChrLis)
    If XHas_Sfx(A, Mid(SfxChrLis, J, 1), Cmp) Then XHas_SfxChr = True: Exit Function
Next
End Function

Function XHas_Sfx(A, Sfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
XHas_Sfx = StrComp(Right(A, Len(Sfx)), Sfx, Cmp) = 0
End Function

Function XHas_SfxAp_IgnCas(A, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
XHas_SfxAp_IgnCas = XHas_SfxAv(A, Av)
End Function
Function XHas_SfxAp(A, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
XHas_SfxAp = XHas_SfxAv(A, Av, vbBinaryCompare)
End Function

Function XHas_SfxAv(A, SfxAv(), Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
Dim Sfx
For Each Sfx In SfxAv
    If XHas_Sfx(A, Sfx, Cmp) Then XHas_SfxAv = True: Exit Function
Next
End Function

Function Ay_IsAllEle_HasPfx(A$(), Pfx$) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
   If Not XHas_Pfx(CStr(I), Pfx) Then Exit Function
Next
Ay_IsAllEle_HasPfx = True
End Function

