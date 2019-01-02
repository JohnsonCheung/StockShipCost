Attribute VB_Name = "MIde_X_WhXXX"
Option Compare Binary
Option Explicit

Function WhBExprSqp$(BExpr$)
If BExpr = "" Then Exit Function
WhBExprSqp = " Where " & BExpr
End Function

Property Get WhEmpNm() As WhNm
End Property

Property Get WhEmpPjMth() As WhPjMth
End Property

Function WhMd(Optional WhCmpTy$, Optional Nm As WhNm) As WhMd
Dim O As New WhMd
O.InCmpTy = CvWhCmpTy(WhCmpTy)
Set O.Nm = Nm
Set WhMd = O
End Function

Function New_WhMdMth(Optional Md As WhMd, Optional Mth As WhMth) As WhMdMth
Set New_WhMdMth = New WhMdMth
With New_WhMdMth
    Set .Md = Md
    Set .Mth = Mth
End With
End Function

Function WhMdMth_WhMd(A As WhMdMth) As WhMd
If Not IsNothing(A) Then Set WhMdMth_WhMd = A.Md
End Function

Function WhMdMth_WhMth(A As WhMdMth) As WhMth
If Not IsNothing(A) Then Set WhMdMth_WhMth = A.Mth
End Function

Function WhMth_Pub_PFX(Pfx$) As WhMth
Set WhMth_Pub_PFX = WhMth("Pub", WhNm_PFX(Pfx))
End Function

Function WhMth_Pub() As WhMth
Set WhMth_Pub = WhMth("Pub")
End Function

Function WhMth(Optional WhMdy$, Optional WhKd$, Optional Nm As WhNm) As WhMth
Set WhMth = New WhMth
With WhMth
    .InShtKd = CvWhKd(WhKd)
    .InMdy = CvWhMdy(WhMdy)
    Set .Nm = Nm
End With
End Function
Function WhPjMth_WhMdMth(A As WhPjMth) As WhMdMth
If IsNothing(A) Then Exit Function
Set WhPjMth_WhMdMth = A.MdMth
End Function

Function WhPjMth_WhPj(A As WhPjMth) As WhNm
If IsNothing(A) Then Exit Function
Set WhPjMth_WhPj = A.Pj
End Function
Function New_WhPjMth(Optional Pj As WhNm, Optional MdMth As WhMdMth) As WhPjMth
Set New_WhPjMth = New WhPjMth
With New_WhPjMth
    Set .Pj = Pj
    Set .MdMth = MdMth
End With
End Function

Function WhPjMth_XAdd_Pub(A As WhPjMth) As WhPjMth
Dim O As WhPjMth
If IsNothing(A) Then
    Set O = New WhPjMth
Else
    Set O = A
End If
With O
    If IsNothing(.MdMth) Then
        Set .MdMth = New WhMdMth
    End If
    If IsNothing(.MdMth.Mth) Then
        Set .MdMth.Mth = New WhMth
    End If
    If Not Ay_XHas(.MdMth.Mth.InMdy, "Pub") Then
        .MdMth.Mth.InMdy = CvSy(Ay_XAdd_Itm(.MdMth.Mth.InMdy, "Pub"))
    End If
End With
Set WhPjMth_XAdd_Pub = O
End Function

Function WhPjMth_MdMth(A As WhPjMth) As WhMdMth
If IsNothing(A) Then Exit Function
Set WhPjMth_MdMth = A.MdMth
End Function

Function WhPjMth_WhNm(A As WhPjMth) As WhNm
If IsNothing(A) Then Exit Function
Set WhPjMth_WhNm = A.Pj
End Function


Function CvWhCmpTy(WhCmpTy$) As vbext_ComponentType()
Dim O() As vbext_ComponentType, I
For Each I In AyNz(Ssl_Sy(WhCmpTy))
    PushI O, CmpTy_ShtNmTy(I)
Next
CvWhCmpTy = O
End Function

Function CvWhMdy(WhMdy$) As String()
If WhMdy = "" Then Exit Function
Dim O$(), M
O = Ssl_Sy(WhMdy): CvWhMdy1 O
If Ay_XHas(O, "Pub") Then Push O, ""
CvWhMdy = O
End Function
Private Function CvWhMdy1(A$())
Dim M
For Each M In A
    If Not Ay_XHas(ShtMdyAy, M) Then Stop
Next
End Function

Function CvWhKd(WhMthKd$) As String()
If WhMthKd = "" Then Exit Function
Dim O$(), K
O = Ssl_Sy(WhMthKd)
For Each K In O
    If Not Ay_XHas(MthKdAy, K) Then Stop
Next
CvWhKd = O
End Function
