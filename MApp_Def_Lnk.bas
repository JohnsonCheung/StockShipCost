Attribute VB_Name = "MApp_Def_Lnk"
Option Compare Binary
Option Explicit
Const SchmLines$ = _
           "Tbl     $Lnk    *Id | InpNm     | FilTy Ffn WhBExpr" _
& vbCrLf & "Tbl     $LnkFld     | InpNm Fld | ExtNm DaoMulTyStr" _
& vbCrLf & "Tbl     $LnkFilTy   | FilTy     | FilTyDes" _
& vbCrLf & "Fld*   *Id InpNm    | ExtNm DaoMulTyStr" _
& vbCrLf & "Fld    $LnkFld    *Id InpNm | ExtNm DaoMulTyStr" _
& vbCrLf & "TblVal $LnkFilTy 1 [aaaa]"

Sub EdtLnk()
With Access.Application
    .Visible = True
    .DoCmd.OpenTable "$Lnk"
End With
End Sub

Sub XEnsLnkDef()
Static A$
If CurDb.Name <> A Then
    'SchmXEns Schm
End If
End Sub

Private Sub Z_XEnsLnkDef()
XEnsLnkDef
End Sub

Private Sub Z()
Z_XEnsLnkDef
MApp_Def_Lnk:
End Sub
