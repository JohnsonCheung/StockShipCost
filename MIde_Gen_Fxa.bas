Attribute VB_Name = "MIde_Gen_Fxa"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Gen_Fxa."
Private FxaWbAy() As Workbook 'Any openned FxaWb will have an entry in this array (by Fxa_XOpn).  So that Fxa can be closed
Type Dst
    PjAy() As VBProject
End Type
Sub Pj_XDmp_Dst(A As VBProject)
Dst_XDmp Pj_Dst(A)
End Sub
Sub CurDst_XDmp()
CurPj_XDmp_Dst
End Sub
Sub CurPj_XDmp_Dst()
Pj_XDmp_Dst CurPj
End Sub
Sub Dst_XDmp(A As Dst)
Dim P
For Each P In AyNz(A.PjAy)
    Debug.Print CvPj(P).Name
Next
End Sub

Sub Pj_DstRfDrs_XDmp(A As VBProject)
Drs_XDmp Pj_DstRfDrs(A)
End Sub
Sub CurPj_DstRfDrs_XDmp()
Pj_DstRfDrs_XDmp CurPj
End Sub

Function Pj_DstRfDrs(A As VBProject) As Drs
Set Pj_DstRfDrs = Dst_RfDrs(Pj_Dst(A))
End Function

Function Pj_Dst(A As VBProject) As Dst
Pj_Dst = XlsPj_Dst(CurXls, A)
End Function

Function Dst_RfDrs(A As Dst) As Drs
Set Dst_RfDrs = PjAyRfDrs(A.PjAy)
End Function

Function XlsPj_Dst(A As excel.Application, Pj As VBProject) As Dst
Dim FxaNy$(), N
For Each N In AyNz(Pj_FxaNy(Pj))
    PushObj XlsPj_Dst.PjAy, XlsPjFxaNm_XOpn_Pj(A, Pj, N)
Next
End Function

Private Function PjFxaNm_FxaNmClsAy_OWNING(A As VBProject, FxaNm) As VBComponent()
Dim ClsNm
For Each ClsNm In AyNz(FxaNm_ClsNy_OWNING(FxaNm))
    PushObj PjFxaNm_FxaNmClsAy_OWNING, A.VBComponents(ClsNm)
Next
End Function
Private Sub XCpy_Cls(FmPj As VBProject, ToPj As VBProject, FxaNm)
CmpAy_XCpy PjFxaNm_FxaNmClsAy_OWNING(FmPj, FxaNm), ToPj
End Sub

Private Sub XCpy_Mod(FmPj As VBProject, ToPj As VBProject, FxaNm)
Dim I
For Each I In AyNz(PjModAy_PFX(FmPj, FxaNm))
    'TimBeg Md_Nm(CvMd(I))
    Md_XCpy CvMd(I), ToPj
Next
'TimEnd
End Sub

Sub CurPj_XBld_DstFxa()
Pj_XBld_Fxa CurPj
End Sub

Function CurPj_FxaNm_Fxa$(FxaNm)
CurPj_FxaNm_Fxa = PjFxaNm_Fxa(CurPj, FxaNm)
End Function
Function CurPj_FxaNm_XOpn_Pj(FxaNm) As VBProject
Set CurPj_FxaNm_XOpn_Pj = XlsPjFxaNm_XOpn_Pj(CurXls, CurPj, FxaNm)
End Function

Sub CurXls_XOpn_Fxa(Fxa)
XlsFxa_XOpn CurXls, Fxa
End Sub

Private Function Pj_ClsAy_OWNING(A As VBProject) As CodeModule()
Dim N
For Each N In AyNz(PjNm_ClsNy_OWNING(Pj_Nm(A)))
    PushObj Pj_ClsAy_OWNING, Pj_Md(A, N)
Next
End Function

Private Function RfNm_GUIDLin$(RfNm)
Const CSub$ = CMod & "RfNm_GUIDLin"
Dim D As Dictionary
Set D = RfDfn_STD
If D.Exists(RfNm) Then RfNm_GUIDLin = D(RfNm): Exit Function
XThw CSub, "Given RfNm cannot find the STD GUID Dic", "RfNm RfDfn_STD", RfNm, Dic_Fmt(D)
End Function

Private Function PjNm_RfGUIDLinAy(PjNm$) As String()
Dim RfNm
For Each RfNm In AyNz(WPjRfNy_STD(PjNm))
    PushI PjNm_RfGUIDLinAy, RfNm & " " & RfNm_GUIDLin(RfNm)
Next
End Function

Private Function WPjRfNy_STD(PjNm$) As String()
Const CSub$ = CMod & "WPjRfNy_STD"
Dim D As Dictionary
Set D = PjRfDfn_STD
If Not D.Exists(PjNm) Then
    XThw CSub, "Given Pj not found in RfDfn_STD", "Pj RfDfn_STD", PjNm, Dic_Fmt(PjRfDfn_STD)
End If
WPjRfNy_STD = Ssl_Sy(D(PjNm))
End Function

Function FxaNm_Fxa$(FxaNm)
FxaNm_Fxa = CurPj_FxaNm_Fxa(FxaNm)
End Function

Sub FxaNm_XAss(A As VBProject, FxaNm, Fun$)
Dim Ny$(): Ny = Pj_FxaNy(A)
If Ay_XHas(Ny, FxaNm) Then Exit Sub
XDmp_Ly Fun, "FxaNm is not found in Pj", "[Invalid FxaNm] Pj [Valid FxaNy]", FxaNm, Pj_Nm(A), Ny
Stop
End Sub

Sub FxaNm_XBld(FxaNm)
PjFxaNm_XBld_Fxa CurPj, FxaNm
End Sub

Sub FxaNm_XOpn(FxaNm)
Fxa_XOpn FxaNm_Fxa(FxaNm)
End Sub

Sub FxaNmVis(FxaNm)
FxaNm_XOpn FxaNm
CurXls.Visible = True
End Sub

Function FxaNm_XOpn_Pj(FxaNm) As VBProject
Dim A$
    A = FxaNm_Fxa(FxaNm)
Fxa_XOpn A
Set FxaNm_XOpn_Pj = VbePj_Ffn(CurXls.Vbe, A)
End Function

Sub Fxa_XOpn(Fxa)
CurXls_XOpn_Fxa Fxa
End Sub

Property Get FxaPth$()
FxaPth = Pj_FxaPth(CurPj)
End Property

Sub FxaPth_XBrw()
Pth_XBrw FxaPth
End Sub

Sub Pj_XBld_Fxa(A As VBProject) ' Build Qxxx.xlsa files under VbePth\Dist\Fxa\*
Dim I, X As excel.Application
Set X = New_Xls
For Each I In Pj_FxaNm_InDependOrder_ByPj(A)
    PjFxaNm_XBld_Fxa A, I, X
Next
Xls_XQuit X
End Sub

Function PjFxaNm_Fxa$(A As VBProject, FxaNm)
PjFxaNm_Fxa = PjFxaNm_Fxa(A, FxaNm)
End Function

Sub PjFxaNm_XBld_Fxa(A As VBProject, FxaNm, Optional Xls As excel.Application)
Const CSub$ = CMod & "PjFxaNm_XBld_Fxa"
FxaNm_XAss A, FxaNm, CSub
Dim X As excel.Application
Dim Fxa$
Dim TarPj As VBProject
    If IsNothing(Xls) Then
        Set X = New_Xls
    Else
        Set X = Xls
    End If
    Fxa = PjFxaNm_Fxa(A, FxaNm)
Pth_XEnsAll Ffn_Pth(Fxa)       '<==Create Dist Pth
Ffn_XDltIfExist Fxa           '<==Create blank new Fxa
Ffn_NotExistAss Fxa, CSub
'TimBeg "Creating Fxa"
Set TarPj = XlsFxa_XCrt(X, Fxa) '<== Open as TarPj
'TimBeg "XCpy_Mod"
XCpy_Mod A, TarPj, FxaNm          '<== Copy module
'TimEnd True
XCpy_Cls A, TarPj, FxaNm           '<== Copy Class
Pj_XSet_Rf_ByRfFfn_STD TarPj      '<== Add Rf
Pj_XSet_Rf_ByRfFfn_USR TarPj     '<== Add Rf

Pj_XSav TarPj                     '<== Sav

Debug.Print Pj_Nm(TarPj); " done ....."

If IsNothing(Xls) Then Xls_XQuit X
End Sub
Function FxaNm_Pj(FxaNm) As VBProject
Set FxaNm_Pj = CurPj_FxaNm_XOpn_Pj(FxaNm)
End Function
Function PjFxaNm_XOpn_Pj(A As VBProject, FxaNm) As VBProject
Set PjFxaNm_XOpn_Pj = XlsPjFxaNm_XOpn_Pj(CurXls, A, FxaNm)
End Function

Sub XlsPjFxa_XOpn(A As excel.Application, Pj As VBProject, Fxa)
If Not Vbe_XHas_Pj_Ffn(A.Vbe, Fxa) Then
    A.Workbooks.Open Fxa
End If
End Sub
Function XlsPjFxaNm_XOpn_Pj(A As excel.Application, Pj As VBProject, FxaNm) As VBProject
Dim Fxa$
    Fxa = PjFxaNm_Fxa(Pj, FxaNm)
XlsPjFxa_XOpn A, Pj, Fxa
Set XlsPjFxaNm_XOpn_Pj = VbePj_Ffn(A.Vbe, Fxa)
End Function
Property Get CurPj_FxaNy() As String()
CurPj_FxaNy = Pj_FxaNy(CurPj)
End Property
Property Get FxaNy() As String()
FxaNy = CurPj_FxaNy
End Property
Function Pj_FxaNy(A As VBProject) As String()
Const CSub$ = CMod & "Pj_FxaNy"
Pj_FxaNy = Ay_XWh_Dist(Ay_XTak_BefOrAll(PjModNy(A), "_"))
AyDupAss Pj_FxaNy, CSub, IgnCas:=True ' If there are 2 Fxa with same name (IgnCas), throw error
End Function

Private Function Pj_FxaNm_InDependOrder_ByPj(A As VBProject) As String()
Pj_FxaNm_InDependOrder_ByPj = Ay_XSrt__BY_AY(Pj_FxaNy(A), FxaNm_InDependOrder)
End Function

Function Pj_FxaPth$(A As VBProject)
Static X$, Y As VBProject
If ObjPtr(A) <> ObjPtr(Y) And X = "" Then X = Pj_Pth(A) & "Dist\Fxa\": Pth_XEns X
Pj_FxaPth = X
End Function

Function PjNm_ClsNy_OWNING(PjNm) As String()
PjNm_ClsNy_OWNING = Ssl_Sy(PjOwnCls(PjNm))
End Function

Function FxaNm_ClsNy_OWNING(FxaNm) As String()
FxaNm_ClsNy_OWNING = PjNm_ClsNy_OWNING(FxaNm)
End Function

Sub FxaNm_XCls(FxaNm)
Dim Fxa1$
Fxa1 = FxaNm_Fxa(FxaNm)
Dim W
For Each W In FxaWbAy
    If Wb_FullNm(CvWb(W)) = Fxa1 Then CvWb(W).Close: ClrFxaWbAy: Exit Sub
Next
End Sub

Private Sub ClrFxaWbAy()
Dim O() As Workbook, Fnd As Boolean, W
For Each W In AyNz(FxaWbAy)
    If Wb_FullNm(CvWb(W)) <> "" Then
        PushObj O, W
        Fnd = True
    End If
Next
If Fnd Then
    FxaWbAy = O
    Stop
End If
End Sub
Sub XlsFxa_XOpn(A As excel.Application, Fxa)
Const CSub$ = CMod & "XlsFxa_XOpn"
If Not IsFxa(Fxa) Then XThw CSub, "Given File is not Fxa", "Fxa", Fxa
If Vbe_XHas_Pj_Ffn(A.Vbe, Fxa) Then Exit Sub
If A.Workbooks.Count = 0 Then A.Workbooks.Add
PushObj FxaWbAy, A.Workbooks.Open(Fxa)
End Sub

Function XlsFxa_XOpn_Pj(A As excel.Application, Fxa) As VBProject

XlsFxa_XOpn A, Fxa
Set XlsFxa_XOpn_Pj = XlsFxa_Pj(A, Fxa)
End Function

Private Sub ZZ_AddRf_STD()
Dim F$, P As VBProject
    F = TmpPth & "MVb.xlam"
    Set P = Fxa_XCrt(F)
Pj_XSet_Rf_ByRfFfn_STD P
Fxa_XOpn F
Stop
End Sub

Private Sub Z_PjNm_RfGUIDLinAy()
Dim PjNm$
PjNm = "MAdoX"
Ept = Ap_Sy( _
    "{420B2830-E718-11CF-893D-00A0C9054228} 1  0 C:\Windows\SysWOW64\scrrun.dll", _
    "{00000600-0000-0010-8000-00AA006D2EA4} 6  0 C:\Program Files (x86)\Common Files\System\ado\msadox.dll")
GoSub Tst
Exit Sub
Tst:
    Act = PjNm_RfGUIDLinAy(PjNm)
    Stop
    C
    Return
End Sub

Private Sub Z_Pj_XBld_Fxa()
Pj_XBld_Fxa CurPj
End Sub

Private Sub Z_PjFxaNm_XBld_Fxa()
Dim Pj As VBProject, FxaNm$
GoSub T1
Exit Sub
T1:
    Set Pj = CurPj
    FxaNm = "MVb"
    GoSub Tst
    Return
Tst:
    Dim Fxa$
    Fxa = PjFxaNm_Fxa(Pj, FxaNm)
    Ffn_XDltIfExist Fxa
    PjFxaNm_XBld_Fxa Pj, FxaNm
    If Not Ffn_Exist(Fxa) Then Stop
    Return
End Sub

Private Sub Z_Pj_FxaNy()
Brw Pj_FxaNy(CurPj)
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As VBProject
Dim C$
Dim D As excel.Application
FxaNm_XBld A
CurPj_XBld_DstFxa
CurPj_FxaNm_XOpn_Pj A
CurXls_XOpn_Fxa A
FxaNm_XAss B, A, C
FxaNm_XBld A
FxaNm_XOpn A
Fxa_XOpn A
FxaPth_XBrw
Pj_XBld_Fxa B
PjFxaNm_Fxa B, A
PjFxaNm_XBld_Fxa B, A, D
Pj_FxaNy B
Pj_FxaPth B
PjNm_ClsNy_OWNING C
XlsFxa_XOpn D, A
XlsFxa_XOpn_Pj D, A
End Sub

Private Sub Z()
Z_PjNm_RfGUIDLinAy
Z_Pj_XBld_Fxa
Z_PjFxaNm_XBld_Fxa
Z_Pj_FxaNy
End Sub
