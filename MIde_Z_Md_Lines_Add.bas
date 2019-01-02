Attribute VB_Name = "MIde_Z_Md_Lines_Add"
Option Compare Binary
Option Explicit
Const CMod$ = ""
Sub Md_XIns_DclLin(A As CodeModule, DclLines$)
A.InsertLines A.CountOfDeclarationLines + 1, DclLines
Debug.Print QQ_Fmt("Md_XIns_DclLin: Module(?) a DclLin is inserted", Md_Nm(A))
End Sub

Sub Md_XApp_Lines(A As CodeModule, Lines$)
Const CSub$ = CMod & "Md_XApp_Lines"
If Lines = "" Then Exit Sub
Dim Bef&, Aft&, Exp&, Cnt&
Bef = A.CountOfLines
A.InsertLines A.CountOfLines + 1, Lines '<=====
Aft = A.CountOfLines
Cnt = Lines_LinCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
'    XThw CSub, "After copy line count are inconsistents, where [Md], [LinCnt-Bef-Cpy], [LinCnt-of-lines], [Exp-LinCnt-Aft-Cpy], [Act-LinCnt-Aft-Cpy], [Lines]", _
        Md_Nm(A), Bef, Cnt, Exp, Aft, Lines
End If
End Sub

Sub Md_XAppy_Ly(A As CodeModule, Ly$())
Md_XApp_Lines A, JnCrLf(Ly)
End Sub

