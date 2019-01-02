Attribute VB_Name = "MDao_TF_Inf"
Option Compare Binary
Option Explicit

Function TF_DaoTy(T, F) As DAO.DataTypeEnum
TF_DaoTy = Dbtf_DaoTy(CurDb, T, F)
End Function

Function TF_DaoShtTyStr$(T, F)
TF_DaoShtTyStr = DaoTy_DaoShtTyStr(TF_DaoTy(T, F))
End Function


