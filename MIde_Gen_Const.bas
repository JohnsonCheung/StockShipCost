Attribute VB_Name = "MIde_Gen_Const"
Option Compare Binary
Option Explicit
Private A_Nm$
Private B_Pj As VBProject
Private B_Md As CodeModule
Private B_Ft$
Private B_ValFmSrc$

Sub ConstFunNm_XEdt(ConstFunNm$)
ZSetAB ConstFunNm
Str_XWrt B_ValFmSrc, B_Ft, True
Ft_XBrw B_Ft
End Sub

Sub ConstFunNm_XUpd_Src_BY_STR(ConstFunNm$, Str$, Optional IsPub As Boolean)
ZSetAB ConstFunNm
Str_XWrt Str, B_Ft, OvrWrt:=True
ConstFunNm_XUpd_Src ConstFunNm, IsPub
End Sub

Sub ConstFunNm_XUpd_Src(ConstFunNm$, Optional IsPub As Boolean)
ZSetAB ConstFunNm
ConstFunNm_XUpd_Src1 ' Stop if the Mth is not Function XXX$()
MdMthNm_XRmv B_Md, A_Nm
B_Md.InsertLines B_Md.CountOfLines + 1, ConstVal_PrpLines(Ft_Lines(B_Ft), A_Nm, IsPub)
End Sub

Private Sub ConstFunNm_XUpd_Src1()
Dim Lin$
Lin = MdMthNm_MthLin(B_Md, A_Nm): If Lin = "" Then Exit Sub
XShf_MthMdy Lin
If XShf_MthShtTy(Lin) <> "Function" Then Stop: GoTo X
If XShf_Nm(Lin) <> A_Nm Then Stop: GoTo X
If XShf_MthShtTyChr(Lin) <> "$" Then Stop: GoTo X
Exit Sub
X: XHalt
End Sub

Private Sub ZSetAB(ConstFunNm$)
A_Nm = ConstFunNm
Dim Pth$
Pth = TmpHom & "GenConst\": Pth_XEns Pth
Set B_Pj = CurPj
Set B_Md = CurMd
B_Ft = Pth & A_Nm & ".txt"
B_ValFmSrc = MthLines_ConstVal(MdMthNm_Lines(B_Md, A_Nm))
End Sub

Private Sub Z_ConstFunNm_XEdt()
ConstFunNm_XEdt "ZZ_A"
End Sub

Private Sub Z_ConstFunNm_XUpd_Src()
ConstFunNm_XUpd_Src "ZZ_A"
End Sub

Private Sub Z()
Z_ConstFunNm_XEdt
Z_ConstFunNm_XUpd_Src
Exit Sub
'ConstFunNm_XEdt
'ConstStrUpdSrc
'ConstFunNm_XUpd_Src
End Sub

Private Property Get Z_Db_XCrt_Schm1$()
Const A_1$ = "Tbl A *Id | *Nm     | *Dte AATy Loc Expr Rmk" & _
vbCrLf & "Tbl B *Id | AId *Nm | *Dte" & _
vbCrLf & "Fld Txt AATy" & _
vbCrLf & "Fld Loc Loc" & _
vbCrLf & "Fld Expr Expr" & _
vbCrLf & "Fld Mem Rmk" & _
vbCrLf & "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']" & _
vbCrLf & "Ele Expr Txt [Expr=Loc & 'abc']" & _
vbCrLf & "Des Tbl     A     AA BB " & _
vbCrLf & "Des Tbl     A     CC DD " & _
vbCrLf & "Des Fld     ANm   AA BB " & _
vbCrLf & "Des Tbl.Fld A.ANm TF_Des-AA-BB"

Z_Db_XCrt_Schm1 = A_1
End Property

