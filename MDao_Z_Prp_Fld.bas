Attribute VB_Name = "MDao_Z_Prp_Fld"
Option Compare Binary
Option Explicit

Function Fd_PrpNy(A As DAO.Field) As String()
Fd_PrpNy = Itr_Ny(A.Properties)
End Function

Property Get TF_Prp(T, F, P)
TF_Prp = Dbtf_Prp(CurDb, T, F, P)
End Property

Property Let TF_Prp(T, F, P, V)
Dbtf_Prp(CurDb, T, F, P) = V
End Property

Private Sub Z_TF_Prp()
XRfh_TmpTbl
Dim P$
P = "Ele"
Ept = 123
GoSub Tst
Exit Sub
Tst:
    TF_Prp("Tmp", "F1", P) = Ept
    Act = TF_Prp("Tmp", "F1", P)
    C
    Return
End Sub

Function Fd_Des$(A As DAO.Field)
If Prps_XHas_Prp(A.Properties, C_Des) Then Fd_Des = A.Properties(C_Des)
End Function


Property Get Dbtf_Des$(A As Database, T, F)
Dbtf_Des = Dbtf_Prp(A, T, F, C_Des)
End Property

Property Let Dbtf_Des(A As Database, T, F, Des$)
Dbtf_Prp(A, T, F, C_Des) = Des
End Property


Private Sub Z()
Z_TF_Prp
MDao_Z_Prp_Fld:
End Sub
