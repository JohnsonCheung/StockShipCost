Attribute VB_Name = "MApp___Fun"
Option Compare Binary
Option Explicit

Sub XEns()
XEns_MdCSub
XEns_OptExp_MD
XEns_SubZZZ_MD
XSrt
End Sub
Property Get AppDtaHom$()
AppDtaHom = PthUp(TmpHom)
End Property

Property Get AppDtaPth$()
AppDtaPth = Pth_XEns(AppDtaHom & Apn & "\")
End Property


Property Get AppFbAy() As String()
Push AppFbAy, AppJJFb
Push AppFbAy, AppStkShpCstFb
Push AppFbAy, AppStkShpRateFb
Push AppFbAy, AppTaxExpCmpFb
Push AppFbAy, AppTaxRateAlertFb
End Property

Property Get AppMdNy() As String()
AppMdNy = Itr_Ny(CodeProject.AllModules)
End Property

Property Get AppPushAppFcmd$()
AppPushAppFcmd = WPth & "PushApp.Cmd"
End Property

Property Get AppRoot$()
Stop '
End Property

Property Get AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
SpecXEnsTbl

Db_XLnk_Ccm CurDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Property

Sub XBrw_DtaFb()
Acs.OpenCurrentDatabase DtaFb
End Sub

Sub XCrt_DtaFb()
If IsDev Then Exit Sub
If Ffn_Exist(DtaFb) Then Exit Sub
Fb_XCrt DtaFb
Dim Src, Tar$, TarFb$
TarFb = DtaFb
Stop
'For Each Src In CcmTny
    Tar = Mid(Src, 2)
    Application.DoCmd.CopyObject TarFb, Tar, acTable, Src
    Debug.Print MsgAp_Lin("XCrt_DtaFb: Cpy [Src] to [Tar]", Src, Tar)
'Next
End Sub

Sub Doc()
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

Property Get DtaDb() As Database
Set DtaDb = DAO.DBEngine.OpenDatabase(DtaFb)
End Property

Property Get DtaFb$()
DtaFb = AppHom & DtaFn
End Property

Property Get DtaFn$()
DtaFn = Apn & "_Data.accdb"
End Property

Property Get IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not PthIsExist(ProdPth)
End If
IsDev = Y
End Property

Property Get IsProd() As Boolean
IsProd = Not IsDev
End Property

Private Sub Z_AppFbAy()
Dim F
For Each F In AppFbAy
If Not Ffn_Exist(F) Then Stop
Next
End Sub

Property Get ProdPth$()
ProdPth = "N:\SAPAccessReports\"
End Property

Private Sub ZZ()
Dim A As Database
XBrw_DtaFb
XCrt_DtaFb
Doc
XEns
End Sub

Private Sub Z()
Z_AppFbAy
End Sub
