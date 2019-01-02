Attribute VB_Name = "MIde_Z_Md_Op_Add_Mth"
Option Compare Binary
Option Explicit
Sub Md_XAdd_Fun(A As CodeModule, Nm$, Lines)
Md_XAdd_1 A, Nm, Lines, IsFun:=True
End Sub

Private Sub Md_XAdd_1(A As CodeModule, Nm$, Lines, IsFun As Boolean)
Dim L$
    Dim B$
    B = IIf(IsFun, "Function", "Sub")
    L = QQ_Fmt("? ?()|?|End ?", B, Nm, Lines, B)
Md_XApp_Lines A, L
Mth_XGo New_Mth(A, Nm)
End Sub

Sub Md_XAdd_Sub(A As CodeModule, Nm$, Lines)
Md_XAdd_1 A, Nm, Lines, IsFun:=False
End Sub
