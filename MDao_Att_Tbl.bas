Attribute VB_Name = "MDao_Att_Tbl"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Att_Tbl."
Private Const Nm$ = "!Att"

Private Property Get AttTd() As DAO.TableDef
Static X As DAO.TableDef
If IsNothing(X) Then
    Set X = CurDb.TableDefs(Nm)
End If
Set AttTd = X
End Property

Sub XEnsTblAtt()
Dim AttTdStr$, B As StruBase
Stop '
DbTdStr_XEns_B CurDb, AttTdStr, B
End Sub

Private Sub ZZ()
End Sub

Private Sub Z()
End Sub
