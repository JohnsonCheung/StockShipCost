Attribute VB_Name = "MVb_Ay_Ap"
Option Compare Binary
Option Explicit

Function FmTo_IntAy(FmInt%, ToInt%) As Integer()
Dim O%(), I&, V%
ReDim O(ToInt - FmInt)
For V = FmInt To ToInt
    O(I) = V
    I = I + 1
Next
FmTo_IntAy = O
End Function
Function FmTo_LngAy(FmLng&, ToLng&) As Long()
Dim O&(), I&, V&
ReDim O(ToLng - FmLng)
For V = FmLng To ToLng
    O(I) = V
    I = I + 1
Next
FmTo_LngAy = O
End Function

Function Ap_DotLin$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_DotLin = JnDot(Av)
End Function
Function Ay_QQInto_T(A, OInto)
Dim O
O = Ay_XReSz(OInto, A)
Dim J&
For J = 0 To UB(A)
    O(J) = A(J)
Next
Ay_QQInto_T = O
End Function
Function Ay_DteAy(A) As Date()
Ay_DteAy = Ay_QQInto_T(A, Ay_DteAy)
End Function
Function Ap_DteAy(ParamArray Ap()) As Date()
Dim Av(): Av = Ap
Ap_DteAy = Ay_DteAy(Av)
End Function

Function Ap_IntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
Ap_IntAy = AyIntAy(Av)
End Function

Function Ap_JnCrLfNOBLANK$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnCrLfNOBLANK = JnCrLf(Ay_SyNOBLANK(Av))
End Function

Function Ap_JnDblDollar$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnDblDollar = JnDollar(Av)
End Function

Function Ap_JnDollar$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnDollar = JnDollar(Av)
End Function

Function Ap_JnDot$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnDot = JnDot(Av)
End Function

Function Ap_JnPthSep$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnPthSep = JnPthSep(Av)
End Function

Function Ap_JnVBar$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnVBar = JnVBar(Av)
End Function

Function Ap_JnVBarSpc$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnVBarSpc = JnVBarSpc(Av)
End Function

Function Ap_Lin$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_Lin = JnSpc(Ay_XExl_EmpEle(Av))
End Function

Function Ap_Lines$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_Lines = JnCrLf(Ay_XExl_EmpEle(Av))
End Function

Function Ap_LngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
Ap_LngAy = AyLngAy(Av)
End Function

Function Ap_JnSemiColon$(ParamArray Ap())
Dim Av(): Av = Ap
Ap_JnSemiColon = JnSemiColon(Ay_XExl_EmpEle(Av))
End Function

Function Ap_SngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
Ap_SngAy = AySngAy(Av)
End Function
Function Ap_Sy_NOBLANK(ParamArray Itm_or_Ay_Ap()) As String()
Dim Av(): Av = Itm_or_Ay_Ap
Dim I
For Each I In Av
    If IsArray(I) Then
        PushNonBlankSy Ap_Sy_NOBLANK, CvSy(I)
    Else
        PushNonBlankStr Ap_Sy_NOBLANK, I
    End If
Next
End Function
Function Ap_Sy(ParamArray Itm_or_Ay_Ap()) As String()
Dim Av(): Av = Itm_or_Ay_Ap
Dim I
For Each I In Av
    If IsArray(I) Then
        PushIAy Ap_Sy, I
    Else
        PushI Ap_Sy, I
    End If
Next
End Function

Private Sub ZZ()
Dim A()
Dim B As Variant

Ap_Sy A
End Sub

Private Sub Z()
End Sub
