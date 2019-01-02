Attribute VB_Name = "MIde_Mth_Op_Cpy"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Mth_Op_Cpy."
Function Mth_XCpy(A As Mth, ToMd As CodeModule, Optional IsSilent As Boolean) As Boolean
Const CSub$ = CMod & "Mth_XCpy"
If Md_XHas_MthNm(ToMd, A.Nm) Then
    If Not IsSilent Then
'        FunMsgAp_XDmpLin CSub, "[FmMth] is Found in [ToMd]", Mth_MthDNm(A), Md_Nm(ToMd)
    End If
    Mth_XCpy = True
    Exit Function
End If
If ObjPtr(A.Md) = ObjPtr(ToMd) Then
    If Not IsSilent Then
'        FunMsgAp_XDmpLin CSub, "[FmMth] module cannot be same as [ToMd]", Mth_MthDNm(A), Md_Nm(ToMd)
    End If
    Mth_XCpy = True
    Exit Function
End If
Md_XApp_Lines ToMd, vbCrLf & MthLines(A)
If Not IsSilent Then
'    FunMsgAp_XDmpLin CSub, "[FmMth] is copied [ToMd]", Mth_MthDNm(A), Md_Nm(ToMd)
End If
End Function
