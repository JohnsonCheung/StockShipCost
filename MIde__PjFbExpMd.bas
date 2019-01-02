Attribute VB_Name = "MIde__PjFbExpMd"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde__PjFbExpMd."

Private Function PjFbExpMd1(A As VBProject) As VBComponent()
Dim C As VBComponent, I
Dim Ny$()
    Ny = Ssl_Sy("Drs New_TblImpSpec LnkCol")
For Each I In AyNz(PjClsAndModCmpAy(A))
    If Not Ay_XHas(Ny, CvCmp(I).Name) Then
        PushObj PjFbExpMd1, I
    End If
Next
End Function

Function PjFbExpMd$(A As VBProject, Fb$)
Ffn_XAss_Exist Fb, CSub
Dim O$
Dim T As VBProject
Dim C() As VBComponent
    O = Ffn_XCpy_ToNxtFfn(Fb)
    Set T = ZFbPj(O)
    C = PjFbExpMd1(A)
ZAss T, C
ZCpy T, C
ZRmvDupMth T
PjFbExpMd = O
End Function

Private Sub ZAss(TarPj As VBProject, CmpAy() As VBComponent)
Const CSub$ = CMod & "ZAss"
Dim N$()
    Dim N1$(): N1 = Itr_Ny(CmpAy)
    Dim N2$(): N2 = Pj_ClsAndModNy(TarPj)
    N = AyIntersect(N1, N2)
If Sz(N) = 0 Then Exit Sub
XThw CSub, "Modules already exist in TarPj", "[Already existed in tar pj module name] TarPjNm TarPjFile", N, Pj_Nm(TarPj), Pj_Ffn(TarPj)
End Sub

Private Sub ZCpy(TarPj As VBProject, CmpAy() As VBComponent)
Dim I
For Each I In AyNz(CmpAy)
    Cmp_XCpy CvCmp(I), TarPj
Next
End Sub

Private Sub ZFbPj1(A As Access.Application, Fb)
If IsNothing(A.CurrentDb) Then
    A.OpenCurrentDatabase Fb
Else
    If A.CurrentDb.Name = Fb Then Exit Sub
    A.CloseCurrentDatabase
    A.OpenCurrentDatabase Fb
End If
End Sub

Private Function ZFbPj(A) As VBProject
Static X As New Access.Application
ZFbPj1 X, A ' OpnCurDb
Set ZFbPj = X.Vbe.ActiveVBProject
End Function


Private Sub Z_PjFbExpMd()
Dim A As VBProject, Fb$
'
Set A = CurPj
Fb = "C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\StockShipCost (ver 1.0).accdb"
GoSub Tst
Exit Sub
Tst:
    Pth_XBrw Ffn_Pth(Fb)
    Act = PjFbExpMd(A, Fb)
    Debug.Print Act
    Stop
    Return
End Sub

Private Sub ZRmvDupMth(A As VBProject) _
'RmvDup Mth in A.bb_Rpt
Dim Ay() As FmCnt, M As CodeModule
    Set M = Pj_Md(A, "bb_Rpt")
    Ay = ZRmvDupMth1(A)
Md_XRmv_FmCntAy M, Ay
End Sub

Private Function ZRmvDupMth1(A As VBProject) As FmCnt()

End Function




Private Sub Z()
Z_PjFbExpMd
MIde__PjFbExpMd:
End Sub
