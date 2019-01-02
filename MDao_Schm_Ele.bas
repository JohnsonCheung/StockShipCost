Attribute VB_Name = "MDao_Schm_Ele"
Option Compare Binary
Option Explicit
Public Const EleLblss$ = "*Ele *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"

Function FdDic_Fmt(A As Dictionary) As String()
If IsNothing(A) Then PushI FdDic_Fmt, "FDic is *Nothing": Exit Function
Dim K
For Each K In A.Keys
    PushI FdDic_Fmt, K & " " & Fd_Str(A(K))
Next
End Function

Function EleStr_Fd(A) As DAO.Field2
Dim TyStr$, Req As Boolean, AlwZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Expr$
Dim L$: L = A
AyAsg XShf_Val(L, EleLblss), _
    TyStr, Req, AlwZLen, Dft, VTxt, VRul, TxtSz, Expr
Set EleStr_Fd = New_Fd( _
    "", DaoShtTyStr_DaoTy(TyStr), Req, TxtSz, AlwZLen, Expr, Dft, VRul, VTxt)
End Function

Private Sub Z_EleStr_Fd()
Dim A$, Act As DAO.Field2, Ept As DAO.Field2
A = "Int Req AlwZLen Dft=ABC TxtSz=10"
Set Ept = New DAO.Field
With Ept
    .Type = DAO.DataTypeEnum.dbInteger
'    .AllowZeroLength = False
    .DefaultValue = "ABC"
    .Required = True
    .Size = 10
End With
GoSub Tst
Exit Sub
Tst:
    Set Act = EleStr_Fd(A)
    If Not Fd_IsEq(Act, Ept) Then Stop
    Return
End Sub

Sub XX()
Dim A As New DAO.Field
Debug.Print A.Name
End Sub

Private Sub Z()
Z_EleStr_Fd
MDao_Schm_Ele:
End Sub
