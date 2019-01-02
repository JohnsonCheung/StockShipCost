Attribute VB_Name = "MIde_Ty_Mth_ShtTy_Cv"
Option Compare Binary
Option Explicit

Function MthShtTy_MthShtKd$(MthTy$)
Select Case MthTy
Case "Fun", "Sub": MthShtTy_MthShtKd = MthTy
Case "Get", "Let", "Set": MthShtTy_MthShtKd = "Prp"
End Select
End Function

Function IsMthTy(A$) As Boolean
IsMthTy = Ay_XHas(MthTyAy, A)
End Function

Function IsMdy(A$) As Boolean
IsMdy = Ay_XHas(MdyAy, A)
End Function

Function MthMdy_MthShtMdy$(A)
Dim O$
Select Case A
Case "Public": O = "Pub"
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
End Select
MthMdy_MthShtMdy = O
End Function

Function MthTy_MthShtKd$(A)
MthTy_MthShtKd = MthShtTy_MthShtKd(MthTy_MthShtTy(A))
End Function

Function MthTy_MthShtTy$(A)
Dim O$
Select Case A
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
End Select
MthTy_MthShtTy = O
End Function

Function ShtMthKd$(MthKd)
Dim O$
Select Case MthKd
Case "Property": O = "Prp"
Case "Function": O = "Fun"
Case "Sub":      O = "Sub"
End Select
ShtMthKd = O
End Function

