Attribute VB_Name = "MIde_Gen_Sln"
Option Compare Binary
Option Explicit
Private Sub Z_ZPjNy()
D ZPjNy
End Sub

Sub GXEnsln()
Stop
Dim N
ZOupPth_XClr
For Each N In ZPjNy
    ZGenPj N
Next
End Sub
Private Sub ZOupPth_XClr()
Pth_XClr ZOupPth
If Sz(Pth_FnAy(ZOupPth)) > 0 Then Stop
End Sub
Private Sub ZOupPth_XBrw()
Pth_XBrw ZOupPth
End Sub
Private Property Get ZOupPth$()
ZOupPth = TmpFdrPth("QFinalSln")
End Property
Private Function ZCrtPj(Pj_Nm) As VBProject
Set ZCrtPj = XlsFxa_Pj(CurXls, ZOupPth & Pj_Nm & ".xlam")
End Function

Private Sub ZGenPj(Pj_Nm)
Dim M, ToPj As VBProject
Set ToPj = ZCrtPj(Pj_Nm)
For Each M In ZPj_MdNy(Pj_Nm)
    Md_XCpy ZSrcMd(M), ToPj
Next
Pj_XSav ToPj
End Sub
Private Function ZSrcMd(MdNm) As CodeModule
Set ZSrcMd = ZSrcPj.VBComponents(MdNm).CodeModule
End Function
Private Property Get ZPjNy() As String()
Dim N
For Each N In PjModNy(ZSrcPj)
    PushNoDupNonBlankStr ZPjNy, ZMdNmPj_Nm(N)
Next
End Property
Private Function ZMdNmPj_Nm__Len%(MdNm)
Dim J%, C$
For J = 5 To Len(MdNm)
    C = Mid(MdNm, J, 1)
    If Asc_IsUCase(Asc(C)) Or C = "_" Then
        ZMdNmPj_Nm__Len = J - 4
        Exit Function
    End If
Next
ZMdNmPj_Nm__Len = J - 4
End Function
Private Function ZMdNmPj_Nm$(MdNm)
If Left(MdNm, 3) <> "Lib" Then Exit Function
Dim L%
    L = ZMdNmPj_Nm__Len(MdNm)
    
ZMdNmPj_Nm = "Q" & Mid(MdNm, 4, L)
End Function
Private Sub Z_ZPj_MdNy()
D ZPj_MdNy("QVb")
End Sub
Private Function ZPj_MdNy(Pj_Nm) As String()
Dim N
For Each N In PjModNy(ZSrcPj)
    If ZMdNmPj_Nm(N) = Pj_Nm Then
        PushNoDup ZPj_MdNy, N
    End If
Next
For Each N In Pj_ClsNy(ZSrcPj)
    If Not ZClsPjDic.Exists(N) Then Stop
    If ZClsPjDic(N) = Pj_Nm Then
        PushNoDup ZPj_MdNy, N
    End If
Next
End Function
Private Property Get ZClsPjLy() As String()
Dim O$()
PushI O, "Blk       QTp"
PushI O, "DCRslt    QVb"
PushI O, "Drs       QDta"
PushI O, "Ds        QDta"
PushI O, "Dt        QDta"
PushI O, "FmCnt     QVb"
PushI O, "FTIx      QVb"
PushI O, "FTNo      QVb"
PushI O, "Gp        QTp"
PushI O, "LCC       QVb"
PushI O, "LnkCol    QDao"
PushI O, "FmCnt    QVb"
PushI O, "Lnx       QTp"
PushI O, "VbeLoc    QIde"
PushI O, "Mth       QIde"
PushI O, "RRCC      QIde"
PushI O, "S1S2      QVb"
PushI O, "New_TblImpSpec    QDao"
PushI O, "WhMd      QIde"
PushI O, "WhMdMth   QIde"
PushI O, "WhMth     QIde"
PushI O, "WhNm      QIde"
PushI O, "WhPjMth   QIde"
PushI O, "P123      QVb"
PushI O, "SwBrk     QTp"
PushI O, "Sql_Shared    QTp"
ZClsPjLy = O
End Property
Private Property Get ZClsPjDic() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = New_Dic_LY(ZClsPjLy)
Set ZClsPjDic = X
End Property
Private Property Get ZSrcPj() As VBProject
Static X As VBProject
If IsNothing(X) Then Set X = Pj("QFinal1")
Set ZSrcPj = X
End Property


Private Sub Z()
Z_ZPj_MdNy
Z_ZPjNy
MIde_Gen_Sln:
End Sub
