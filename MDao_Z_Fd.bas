Attribute VB_Name = "MDao_Z_Fd"
Option Compare Binary
Option Explicit
Function FdClone(A As DAO.Field2, FldNm) As DAO.Field2
Set FdClone = New DAO.Field
With FdClone
    .Name = FldNm
    .Type = A.Type
    .AllowZeroLength = A.AllowZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function FdEleScl$(A As DAO.Field2)
Dim Rq$, Ty$, TxtSz$, AlwZ$, Rul$, Dft$, VTxt$, Expr$, Des$
Des = Val_XAdd_Lbl(Fd_Des(A), "Des")
Rq = Bool_Txt_IfTrue(A.Required, "Req")
AlwZ = Bool_Txt_IfTrue(A.AllowZeroLength, "AlwZ")
Ty = DaoTy_DaoShtTyStr(A.Type)
If A.Type = DAO.DataTypeEnum.dbText Then TxtSz = Bool_Txt_IfTrue(A.Type = dbText, "TxXTSz=" & A.Size)
Rul = Val_XAdd_Lbl(A.ValidationText, "VTxt")
VTxt = Val_XAdd_Lbl(A.ValidationRule, "VRul")
Expr = Val_XAdd_Lbl(A.Expression, "Expr")
Dft = Val_XAdd_Lbl(A.DefaultValue, "Dft")
FdEleScl = Ap_JnSemiColon(Ty, TxtSz, Rq, AlwZ, Rul, VTxt, Dft, Expr)
End Function

Function FdScl$(A As DAO.Field2)
FdScl = A.Name & ";" & FdEleScl(A)
End Function

Function StdFldFd_Str$(F)
StdFldFd_Str = Fd_Str(New_Fd_STD(F))
End Function

Function Fd_Str$(A As DAO.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = DAO.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = " " & XQuote_SqBkt_IfNeed("Dft=" & A.DefaultValue)
If A.Required Then R = " Req"
If A.AllowZeroLength Then Z = " AlwZLen"
If A.Expression <> "" Then E = " " & XQuote_SqBkt_IfNeed("Expr=" & A.Expression)
If A.ValidationRule <> "" Then VRul = " " & XQuote_SqBkt_IfNeed("VRul=" & A.ValidationRule)
If A.ValidationText <> "" Then VRul = " " & XQuote_SqBkt_IfNeed("VTxt=" & A.ValidationText)
Fd_Str = A.Name & " " & DaoTy_DaoShtTyStr(A.Type) & R & Z & S & VTxt & VRul & D & E
End Function

Function Fd_Val(A As DAO.Field)
Fd_Val = A.Value
End Function

Function Fd_SqlTy$(A As DAO.Field)
Stop '
End Function
Function Fd_IsEq(A As DAO.Field2, B As DAO.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AllowZeroLength <> B.AllowZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
Fd_IsEq = True
End Function
Function CvFd(A) As DAO.Field
Set CvFd = A
End Function

Function CvFd2(A As DAO.Field) As DAO.Field2
Set CvFd2 = A
End Function
