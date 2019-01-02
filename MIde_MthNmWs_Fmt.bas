Attribute VB_Name = "MIde_MthNmWs_Fmt"
Option Explicit

Function MthNmWs_XFmt(A As Worksheet) As Worksheet
Dim Lo As ListObject
    Set Lo = Ws_FstLo(A)
XAdd_Col_HasUnderScore Lo
XAdd_Col_IsPrp Lo
XSrt A
Set MthNmWs_XFmt = A
End Function
Private Sub XSrt(A As Worksheet)

End Sub
Private Sub XAdd_Col_IsPrp(A As ListObject)
Lo_XAdd_ColFormula A, "IsPrp", "=OR([Ty]=""Let"",[Ty]=""Get"",[Ty]=""Let"")"
End Sub


Private Sub XAdd_Col_HasUnderScore(A As ListObject)

End Sub
Function Lo_XAdd_ColFormula(A As ListObject, ColNm$, Formula$)
Lo_XHas_Col_XAss A, ColNm, CSub, "Formula", Formula
End Function
Sub Lo_XHas_Col_XAss(A As ListObject, ColNm$, Fun$, ParamArray NN_ValAp())

End Sub
