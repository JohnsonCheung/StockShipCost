Attribute VB_Name = "MApp_Tp"
Option Compare Binary
Const CMod$ = "MApp_Tp."

Option Explicit
Sub XAdd_TpWc()
Dim Wb As Workbook
Set Wb = TpWb
Wb_XAdd_Wc Wb, WFb, "@Main"
Wb_XAdd_Wc Wb, WFb, "@Rate"
Wb_XAdd_Wc Wb, WFb, "@Sku"
Wb_XAdd_Wc Wb, WFb, "@Repack1"
Wb_XAdd_Wc Wb, WFb, "@Repack2"
Wb_XAdd_Wc Wb, WFb, "@Repack3"
Wb_XAdd_Wc Wb, WFb, "@Repack4"
Wb_XAdd_Wc Wb, WFb, "@Repack5"
Wb_XAdd_Wc Wb, WFb, "@Repack6"
Wb.Close True
End Sub
Property Get TpExist() As Boolean
TpExist = Ffn_Exist(TpFx)
End Property
Sub XExp_Tp()
Att_XExp_ToFfn "Tp", TpFx
End Sub

Property Get TpFnn$()
TpFnn = Apn & "(Template)"
End Property

Property Get TpIdxWs() As Worksheet
Set TpIdxWs = Wb_Ws_Cd(TpWb, "WsIdx")
End Property

Sub TpImp()
Const CSub$ = CMod & "TpImp"
Dim A$
A = TpFx
If Not Ffn_Exist(A) Then
    If True Then
        FunMsgNyAp_XDmp CSub, "Tp not exist, no Import", "TpFx", A
    End If
End If
If Att_IsOld("Tp", A) Then Att_XImp "Tp", A '<== Import
End Sub

Property Get TpMainFbtStr$()
Dim Wb As Workbook, Qt As QueryTable
Set Wb = TpWb
Set Qt = Wb_MainQt(Wb)
TpMainFbtStr = Qt_FbtStr(Qt)
Wb_XQuit Wb
End Property

Property Get TpMainLo() As ListObject
Set TpMainLo = Wb_MainLo(TpWb)
End Property

Property Get TpMainQt() As QueryTable
Set TpMainQt = Wb_MainQt(TpWb)
End Property
'===============================================
Property Get TpFx$()
TpFx = TpPth & Apn & "(Template).xlsx"
End Property

Property Get TpFxm$()
TpFxm = TpPth & Apn & "(Template).xlsm"
End Property
Sub XBrw_Tp()
Fx_XBrw TpFx
End Sub

Property Get TpWs_CdNy() As String()
TpWs_CdNy = Fx_WsCdNy(TpFx)
End Property


'==============================================
Sub XMin_TpLo()
Dim O As Workbook
Set O = TpWb
O.Save
Wb_XVis O
End Sub

Property Get TpPth$()
TpPth = Pth_XEns(CurDb_Pth & "Template\")
End Property

Sub XRfh_Tp()
Wb_XVis Wb_XRfh(TpWb)
End Sub

Sub XRfh_TpWc()
Fx_XRfh TpFx, WFb
End Sub

Property Get TpWb() As Workbook
Set TpWb = Fx_Wb(TpFx)
End Property

Property Get TpWcSy() As String()
Dim W As Workbook, X As excel.Application
Set X = New excel.Application
Set W = X.Workbooks.Open(TpFx)
TpWcSy = Wb_WcSy_OLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Property

Sub Ffn_XCrt_ByTp(Ffn$)
Att_XExp_ToFfn "Tp", Ffn
End Sub

