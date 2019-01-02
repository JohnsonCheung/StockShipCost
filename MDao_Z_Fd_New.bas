Attribute VB_Name = "MDao_Z_Fd_New"
Option Compare Binary
Option Explicit

Private Function FdDic_Fmt(A As Dictionary) As String()
Dim K, O$()
For Each K In A.Keys
    PushI O, K & " " & Fd_Str(CvFd2(A(K)))
Next
FdDic_Fmt = Ay_XAlign_4T(O)
End Function

Function New_Fd(F, Optional Ty As DAO.DataTypeEnum = dbText, Optional Req As Boolean, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional VRul$, Optional VTxt$) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Required = Req
    If Ty <> 0 Then .Type = Ty
    If Ty = dbText Then
        .Size = TxtSz
        .AllowZeroLength = ZLen
    End If
    If Expr <> "" Then
        CvFd2(O).Expression = Expr
    End If
    O.DefaultValue = Dft
End With
Set New_Fd = O
End Function

Function New_Fd_BOOL(F) As DAO.Field2
Set New_Fd_BOOL = New_Fd(F, dbBoolean, True, Dft:="0")
End Function

Function New_Fd_CRTDTE(F) As DAO.Field2
Set New_Fd_CRTDTE = New_Fd(F, dbDate, True, Dft:="Now()")
End Function

Function New_Fd_CUR(F) As DAO.Field2
Set New_Fd_CUR = New_Fd(F, dbCurrency, True, Dft:="0")
End Function

Function New_Fd_DBL(F) As DAO.Field2
Set New_Fd_DBL = New_Fd(F, dbDouble, True, Dft:="0")
End Function

Function New_Fd_DTE(F) As DAO.Field2
Set New_Fd_DTE = New_Fd(F, dbDate, True, Dft:="0")
End Function

Function New_Fd_EF(F, T, EF As EF) As DAO.Field2
Set New_Fd_EF = New_Fd_STD(F, CStr(T)): If IsSomething(New_Fd_EF) Then Exit Function
Dim Ele$
Ele = LikssDicKey(EF.E, F): If Ele = "" Then XThw CSub, "Fld cannot lookup from EF", "T F EDic FDic", T, F, EF.E, EF.F
Set New_Fd_EF = EleNm_Fd(F, Ele): If IsSomething(New_Fd_EF) Then Exit Function
If IsNothing(EF.F) Then
    XThw CSub, "F's Ele is not standard and FDic is nothing.  F's Fd cannot be determined.", "F Ele", F, Ele
End If
If Not EF.F.Exists(Ele) Then XThw CSub, "F's Ele is not found in FDic", "F [F's Ele] FDic", F, Ele, FdDic_Fmt(EF.F)
Set New_Fd_EF = EF.F(Ele)
New_Fd_EF.Name = F
End Function

Function EleNm_Fd(EleNm, F) As DAO.Field2
Dim O As DAO.Field2
Set O = EleNm_Fd_TNNN(F, EleNm): If Not IsNothing(O) Then Set EleNm_Fd = O: Exit Function
Select Case EleNm
Case "Nm":  Set EleNm_Fd = New_Fd_NM(F)
Case "Amt": Set EleNm_Fd = New_Fd_CUR(F): EleNm_Fd.DefaultValue = 0
Case "Txt": Set EleNm_Fd = New_Fd_TXT(F, dbText, True): EleNm_Fd.DefaultValue = """""": EleNm_Fd.AllowZeroLength = True
Case "Dte": Set EleNm_Fd = New_Fd_DTE(F)
Case "Int": Set EleNm_Fd = New_Fd_INT(F)
Case "Lng": Set EleNm_Fd = New_Fd_LNG(F)
Case "Dbl": Set EleNm_Fd = New_Fd_DBL(F)
Case "Sng": Set EleNm_Fd = New_Fd_SNG(F)
Case "Lgc": Set EleNm_Fd = New_Fd_BOOL(F)
Case "Mem": Set EleNm_Fd = New_Fd_MEM(F)
End Select
End Function

Private Function EleNm_Fd_TNNN(F, EleTnnn) As DAO.Field2
If Left(EleTnnn, 1) <> "T" Then Exit Function
Dim A$
A = Mid(EleTnnn, 2)
If CStr(Val(A)) <> A Then Exit Function
Set EleNm_Fd_TNNN = New_Fd(F, dbText, True)
With EleNm_Fd_TNNN
    .Size = A
    .DefaultValue = """"""
    .AllowZeroLength = True
End With
End Function

Function New_Fd_FK(F) As DAO.Field2
Set New_Fd_FK = New DAO.Field
With New_Fd_FK
    .Name = F
    .Type = dbLong
End With
End Function

Function New_Fd_ID(F) As DAO.Field2
If Not XHas_Sfx(F, "Id") Then Stop
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
    .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set New_Fd_ID = O
End Function

Function New_Fd_INT(F) As DAO.Field2
Set New_Fd_INT = New_Fd(F, dbInteger, True, Dft:="0")
End Function

Function New_Fd_LNG(F) As DAO.Field2
Set New_Fd_LNG = New_Fd(F, dbLong, True, Dft:="0")
End Function

Function New_Fd_ATT(F) As DAO.Field2
Set New_Fd_ATT = New_Fd(F, dbAttachment)
End Function

Function New_Fd_MEM(F) As DAO.Field2
Set New_Fd_MEM = New_Fd(F, dbMemo, True, Dft:="""""")
End Function

Function New_Fd_NM(F) As DAO.Field2
If Right(F, 2) <> "Nm" Then Stop
Set New_Fd_NM = New_Fd(F, dbText, True, 50, False)
End Function

Function New_Fd_PK(F) As DAO.Field2
If Right(F, 2) <> "Id" Then Stop
Set New_Fd_PK = New_Fd(F, dbLong, True)
New_Fd_PK.Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
End Function

Function New_Fd_SNG(F) As DAO.Field2
Set New_Fd_SNG = New_Fd(F, dbSingle, True, Dft:="0")
End Function

Function New_Fd_STD(F, Optional T$) As DAO.Field2
Dim R2$, R3$: R2 = Right(F, 2): R3 = Right(F, 3)
Select Case True
Case F = "CrtDte": Set New_Fd_STD = New_Fd_CRTDTE(F)
Case T & "Id" = F: Set New_Fd_STD = New_Fd_PK(F)
Case R2 = "Id":    Set New_Fd_STD = New_Fd_ID(F)
Case R2 = "Ty":    Set New_Fd_STD = New_Fd_TY(F)
Case R2 = "Nm":    Set New_Fd_STD = New_Fd_NM(F)
Case R3 = "Dte":   Set New_Fd_STD = New_Fd_DTE(F)
Case R3 = "Amt":   Set New_Fd_STD = New_Fd_CUR(F)
Case R3 = "Att":   Set New_Fd_STD = New_Fd_ATT(F)
End Select
End Function

Function New_Fd_STR(Fd_Str) As DAO.Field2
Dim J%, F$, L$, T$, Ay$(), Sz As Byte, Des$, Rq As Boolean, Ty As DAO.DataTypeEnum, AlwZ As Boolean, Dft$, VRul$, VTxt$, Expr$, Er$()
L = Fd_Str
F = XShf_T(L)
T = XShf_T(L)
Ty = DaoShtTyStr_DaoTy(T)
SclAsg L, VdtEleSclNmSsl, Rq, AlwZ, Sz, Dft, VRul, VTxt, Des, Expr
Dim O As New DAO.Field
With O
    .Name = F
    .DefaultValue = Dft
    .Required = Rq
    .Type = Ty
    If Ty = DAO.DataTypeEnum.dbText Then
        .Size = IIf(Sz = 0, 255, Sz)
        .AllowZeroLength = AlwZ
    End If
    .ValidationRule = VRul
    .ValidationText = VTxt
End With
Set New_Fd_STR = O
End Function

Function New_Fd_TXT(F, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Set New_Fd_TXT = New_Fd(F, dbText, Req, TxtSz, ZLen, Expr, Dft, VRul, VTxt)
End Function

Function New_Fd_TY(F) As DAO.Field2
Set New_Fd_TY = New_Fd(F, dbText, True, 20, ZLen:=False)
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As DAO.DataTypeEnum
Dim C As Boolean
Dim D As Byte
Dim E$
New_Fd_CRTDTE A
New_Fd_CUR A
New_Fd_DTE A
EleNm_Fd A, A
New_Fd_FK A
New_Fd_ID A
New_Fd_NM A
New_Fd_PK A
New_Fd_STD A, E
New_Fd_STR A
New_Fd_TXT A, D, C, E, E, C, E, E
New_Fd_TY A
End Sub

Private Sub Z()
End Sub

