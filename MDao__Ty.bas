Attribute VB_Name = "MDao__Ty"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao__Ty."

Function DaoSemiColonTyStr_DaoTyAy(A) As DAO.DataTypeEnum()
Dim Ay$(), I
    Ay = SplitSC(A)
For Each I In AyNz(Ay)
    PushI DaoSemiColonTyStr_DaoTyAy, DaoShtTyStr_DaoTy(I)
Next
End Function

Function DaoShtTyStr_DaoTy(A) As DAO.DataTypeEnum
Const CSub$ = CMod & "DaoShtTyStr_DaoTy"
Dim O
Select Case A
Case "Lgc": O = DAO.DataTypeEnum.dbBoolean
Case "Dbl": O = DAO.DataTypeEnum.dbDouble
Case "Txt": O = DAO.DataTypeEnum.dbText
Case "Dte": O = DAO.DataTypeEnum.dbDate
Case "Byt": O = DAO.DataTypeEnum.dbByte
Case "Int": O = DAO.DataTypeEnum.dbInteger
Case "Lng": O = DAO.DataTypeEnum.dbLong
Case "Dec": O = DAO.DataTypeEnum.dbDecimal
Case "Cur": O = DAO.DataTypeEnum.dbCurrency
Case "Sng": O = DAO.DataTypeEnum.dbSingle
Case Else: XThw CSub, "Program Error: Invalid DaoShtTyStr.  Check LnkColVbl definition.", "DaoShtTyStr Valid-DaoShtTyStr", A, "Lgc Dbl Txt Dte Byt Int Lng Dec Cur Sng"
End Select
DaoShtTyStr_DaoTy = O
End Function

Function DaoTy_SimTy(A As DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case A
Case _
   DAO.DataTypeEnum.dbBigInt, _
   DAO.DataTypeEnum.dbByte, _
   DAO.DataTypeEnum.dbCurrency, _
   DAO.DataTypeEnum.dbDecimal, _
   DAO.DataTypeEnum.dbDouble, _
   DAO.DataTypeEnum.dbFloat, _
   DAO.DataTypeEnum.dbInteger, _
   DAO.DataTypeEnum.dbLong, _
   DAO.DataTypeEnum.dbNumeric, _
   DAO.DataTypeEnum.dbSingle
   O = eNbr
Case _
   DAO.DataTypeEnum.dbChar, _
   DAO.DataTypeEnum.dbGUID, _
   DAO.DataTypeEnum.dbMemo, _
   DAO.DataTypeEnum.dbText
   O = eTxt
Case _
   DAO.DataTypeEnum.dbBoolean
   O = eLgc
Case _
   DAO.DataTypeEnum.dbDate, _
   DAO.DataTypeEnum.dbTimeStamp, _
   DAO.DataTypeEnum.dbTime
   O = eDte
Case Else
   O = eOth
End Select
DaoTy_SimTy = O
End Function

Function DaoTy_DaoShtTyStr$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbByte: O = "Byt"
Case DAO.DataTypeEnum.dbLong: O = "Lng"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbDate: O = "Dte"
Case DAO.DataTypeEnum.dbText: O = "Txt"
Case DAO.DataTypeEnum.dbBoolean: O = "Yes"
Case DAO.DataTypeEnum.dbDouble: O = "Dbl"
Case DAO.DataTypeEnum.dbCurrency: O = "Cur"
Case DAO.DataTypeEnum.dbMemo: O = "Mem"
Case DAO.DataTypeEnum.dbAttachment: O = "Att"
Case DAO.DataTypeEnum.dbSingle: O = "Sng"
Case DAO.DataTypeEnum.dbDecimal: O = "Dec"
Case Else: O = "?" & A & "?"
End Select
DaoTy_DaoShtTyStr = O
End Function

Function DaoTy_SqlTyStr$(A As DataTypeEnum, Optional Sz%, Optional Precious%)
Stop '
End Function

Function Val_DaoTy(A) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case VarType(A)
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbString: O = dbText
Case VbVarType.vbDate: O = dbDate
Case Else: Stop
End Select
Val_DaoTy = O
End Function

Private Sub ZZ()
Dim A
Dim B As DAO.DataTypeEnum
Dim C As DataTypeEnum
Dim D%
Dim XX
DaoSemiColonTyStr_DaoTyAy A
DaoShtTyStr_DaoTy A
DaoTy_DaoShtTyStr B
DaoTy_SqlTyStr C, D, D
Val_DaoTy A
End Sub

Private Sub Z()
End Sub
