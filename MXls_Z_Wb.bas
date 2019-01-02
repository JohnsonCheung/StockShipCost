Attribute VB_Name = "MXls_Z_Wb"
Option Compare Binary
Option Explicit

Property Get CurWb() As Workbook
Set CurWb = CurXls.ActiveWorkbook
End Property

Function CvWb(A) As Workbook
Set CvWb = A
End Function

Function WbCn_TxtCn(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set WbCn_TxtCn = A.TextConnection
End Function

Function Wb_FstWs(A As Workbook) As Worksheet
Set Wb_FstWs = A.Sheets(1)
End Function
Function Wb_FstWsNm$(A As Workbook)
Wb_FstWsNm = Wb_FstWs(A).Name
End Function

Function Wb_Fx$(A As Workbook)
Dim F$
F = A.FullName
If F = A.Name Then Exit Function
Wb_Fx = F
End Function

Function Wb_LasWs(A As Workbook) As Worksheet
Set Wb_LasWs = A.Sheets(A.Sheets.Count)
End Function

Function Wb_Lo(A As Workbook, LoNm$) As ListObject
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws_XHas_LoNm(Ws, LoNm) Then Set Wb_Lo = Ws.ListObjects(LoNm): Exit Function
Next
End Function

Function Wb_LoAy(Tp As Workbook) As ListObject()
Dim Ws As Worksheet
For Each Ws In Tp.Sheets
    'PushItr Wb_LoAy, Ws.ListObjects
Next
End Function

Function Wb_LoAy_WhTDashPfx(A As Workbook) As ListObject()
Wb_LoAy_WhTDashPfx = Oy_XWh_NmHasPfx(Wb_LoAy(A), "T_")
End Function

Function Wb_MainLo(A As Workbook) As ListObject
Dim O As Worksheet, Lo As ListObject
Set O = Wb_MainWs(A):              If IsNothing(O) Then Exit Function
Set Wb_MainLo = Ws_Lo(O, "T_Main")
End Function

Function Wb_MainQt(A As Workbook) As QueryTable
Dim Lo As ListObject
Set Lo = Wb_MainLo(A): If IsNothing(A) Then Exit Function
Set Wb_MainQt = Lo.QueryTable
End Function

Function Wb_MainWs(A As Workbook) As Worksheet
Set Wb_MainWs = Wb_Ws_Cd(A, "WsOMain")
End Function

Function Wb_OupLoAy(A As Workbook) As ListObject()
Wb_OupLoAy = Oy_XWh_NmHasPfx(Wb_LoAy(A), "T_")
End Function

Function Wb_Par(A As Workbook) As Workbooks
Set Wb_Par = A.Parent
End Function

Function Wb_PtAy(A As Workbook) As PivotTable()
Dim O() As PivotTable, Ws As Worksheet
For Each Ws In A.Sheets
    PushObjAy O, Ws_PtAy(Ws)
Next
Wb_PtAy = O
End Function

Function Wb_PtNy(A As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushIAy Wb_PtNy, WsPtNy(Ws)
Next
End Function

Function Wb_TxtCn(A As Workbook) As TextConnection
Dim N%: N = Wb_TxtCnCnt(A)
If N <> 1 Then
    Stop
    Exit Function
End If
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then
        Set Wb_TxtCn = C.TextConnection
        Exit Function
    End If
Next
Stop
'XHalt_Impossible CSub
End Function

Function Wb_TxtCnCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then Cnt = Cnt + 1
Next
Wb_TxtCnCnt = Cnt
End Function

Function Wb_TxtCnStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = Wb_TxtCn(A)
If IsNothing(T) Then Exit Function
Wb_TxtCnStr = T.Connection
End Function

Function Wb_WcAy_OLE(A As Workbook) As OLEDBConnection()
'Dim O() As OLEDBConnection, Wc As WorkbookConnection
'For Each Wc In A.Connections
'    PushObjNonNothingObj O, Wc.OLEDBConnection
'Next
'Wb_WcAy_OLE = O
Wb_WcAy_OLE = Oy_XWh_NoNothing(ItrPrp_Into(A.Connections, "OLEDBConnection", Wb_WcAy_OLE))
End Function

Function Wb_WcNy(A As Workbook) As String()
Wb_WcNy = Itr_Ny(A.Connections)
End Function

Function Wb_WcSy_OLE(A As Workbook) As String()
Wb_WcSy_OLE = Oy_PrpSy(Wb_WcAy_OLE(A), "Connection")
End Function

Function Wb_Ws(A As Workbook, WsNm) As Worksheet
Set Wb_Ws = A.Sheets(WsNm)
End Function

Function Wb_WsNy(A As Workbook) As String()
Wb_WsNy = Itr_Ny(A.Sheets)
End Function

Function Wb_Ws_BY_CD_NM(A As Workbook, CdNm$) As Worksheet
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.CodeName = CdNm Then Set Wb_Ws_BY_CD_NM = Ws: Exit Function
Next
End Function

Function Wb_Ws_Cd(A As Workbook, Ws_CdNm$) As Worksheet
Set Wb_Ws_Cd = Itr_FstItm_PrpEqV(A.Sheets, "CodeName", Ws_CdNm)
End Function

Function Wb_Ws_CdNy(A As Workbook) As String()
Wb_Ws_CdNy = ItrPrp_Sy(A.Sheets, "CodeName")
End Function

Function Wb_XAdd_Dbt(A As Workbook, Db As Database, T, Optional UseWc As Boolean) As Workbook
Set Wb_XAdd_Dbt = A
End Function
Function Wb_FullNm$(A As Workbook)
On Error Resume Next
Wb_FullNm = A.FullName
End Function

Function Wb_XAdd_Dbtt(A As Workbook, Db As Database, TT, Optional UseWc As Boolean) As Workbook
Ay_XDoPPXP CvNy(TT), "Wb_XAdd_Dbt", A, Db, UseWc
Set Wb_XAdd_Dbtt = A
End Function

Function Wb_XAdd_Dt(A As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = Wb_XAdd_Ws(A, Dt.DtNm)
Drs_Lo Dt_Drs(Dt), Ws_A1(O)
Set Wb_XAdd_Dt = O
End Function

Function Wb_XAdd_Wc(A As Workbook, Fb$, Nm$) As WorkbookConnection
Set Wb_XAdd_Wc = A.Connections.Add2(Nm, Nm, Fb_CnStr_WbCn(Fb), Nm, XlCmdType.xlCmdTable)
End Function

Function Wb_XAdd_Ws(A As Workbook, Optional WsNm, Optional AtBeg As Boolean, Optional AtEnd As Boolean, Optional BefWsNm$, Optional AftWsNm$) As Worksheet
Dim O As Worksheet
Wb_XDlt_Ws A, WsNm
Select Case True
Case AtBeg:         Set O = A.Sheets.Add(Wb_FstWs(A))
Case AtEnd:         Set O = A.Sheets.Add(Wb_LasWs(A))
Case BefWsNm <> "": Set O = A.Sheets.Add(A.Sheets(BefWsNm))
Case AftWsNm <> "": Set O = A.Sheets.Add(, A.Sheets(AftWsNm))
Case Else:          Set O = A.Sheets.Add
End Select
Set Wb_XAdd_Ws = Ws_XSet_Nm(O, WsNm)
End Function

Sub Wb_XHas_OupNy_XAss(A As Workbook, OupNy$())
Dim O$(), N$, B$(), Ws_CdNy$()
Ws_CdNy = Wb_Ws_CdNy(A)
O = AyMinus(Ay_XAdd_Pfx(OupNy, "WsO"), Ws_CdNy)
If Sz(O) > 0 Then
    N = "OupNy":  B = OupNy:  GoSub Dmp
    N = "WbCdNy": B = Ws_CdNy: GoSub Dmp
    N = "Mssing": B = O:      GoSub Dmp
    Stop
    Exit Sub
End If
Exit Sub
Dmp:
Debug.Print UnderLin(N)
Debug.Print N
Debug.Print UnderLin(N)
Ay_XDmp B
Return
End Sub

Sub Wb_XCls_NoSav(A As Workbook)
A.Close False
End Sub

Sub Wb_XDlt_Wc(A As Workbook)
Itr_XDo A.Connections, "Wc_XDlt"
End Sub
Sub Wb_XDlt_FstWs(A As Workbook)
Wb_XDlt_Ws A, Wb_FstWs(A)
End Sub

Sub Wb_XDlt_Ws(A As Workbook, WsNm)
If Wb_XHas_Ws(A, WsNm) Then
    A.Application.DisplayAlerts = False
    Wb_Ws(A, WsNm).Delete
    A.Application.DisplayAlerts = True
End If
End Sub

Sub Wb_XFmt(A As Workbook, WbFmtrFunAv())
Dim I
For Each I In AyNz(WbFmtrFunAv)
    Run I, A
Next
Wb_XMax(Wb_XVis(A)).Save
End Sub

Sub Wb_XFmt_AllLo(A As Workbook)
FmtSpec_XImp
'Ay_XBrwThw FmtSpec_ErLy
Ay_XDoXP Wb_LoAy(A), "Lo_Fmt", FmtSpec_Ly
End Sub

Function Wb_XMax(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set Wb_XMax = A
End Function

Function Wb_XNew_A1(A As Workbook, Optional WsNm$) As Range
Set Wb_XNew_A1 = Ws_A1(Wb_XAdd_Ws(A, WsNm))
End Function

Sub Wb_XQuit(A As Workbook)
Xls_XQuit A.Application
End Sub

Function Wb_XSav(A As Workbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.Save
A.Application.DisplayAlerts = Y
Set Wb_XSav = A
End Function

Function Wb_XSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set Wb_XSavAs = A
End Function

Sub Wb_XSet_WcTxtCn(A As Workbook, Fcsv$)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = Wb_TxtCn(A)
Dim C$: C = T.Connection: If Not XHas_Pfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Function Wb_XVis(A As Workbook) As Workbook
Xls_XVis A.Application
Set Wb_XVis = A
End Function

Function Wb_XHas_Ws(A As Workbook, WsNm) As Boolean
Wb_XHas_Ws = Itr_XHas_Nm(A.Sheets, WsNm)
End Function

Private Sub ZZ_WbWcSy()
'D Wb_WcSy_OLE(Fx_Wb(TpFx))
End Sub

Private Sub ZZ_Wb_LoAy()
'D OyNy(Wb_LoAy(TpWb))
End Sub

Private Sub ZZ_Wb_LoAy_WhTDashPfx()
D Itr_Ny(Wb_LoAy_WhTDashPfx(TpWb))
End Sub

Private Sub Z_Wb_TxtCnCnt()
Dim O As Workbook: 'Set O = Fx_Wb(Vbe_MthFx)
Ass Wb_TxtCnCnt(O) = 1
O.Close
End Sub

Private Sub Z_Wb_XSet_WcTxtCn()
Dim Wb As Workbook
'Set Wb = Fx_Wb(Vbe_MthFx)
Debug.Print Wb_TxtCnStr(Wb)
Wb_XSet_WcTxtCn Wb, "C:\ABC.CSV"
Ass Wb_TxtCnStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub

Private Sub ZZ()
Dim A
Dim B As WorkbookConnection
Dim C As Workbook
Dim D$
Dim E As Database
Dim F As Boolean
Dim G As Dt
Dim H$()
Dim I()
Dim XX
CvWb A
WbCn_TxtCn B
Wb_FstWs C
Wb_Fx C
Wb_LasWs C
Wb_Lo C, D
Wb_LoAy C
Wb_LoAy_WhTDashPfx C
Wb_MainLo C
Wb_MainQt C
Wb_MainWs C
Wb_OupLoAy C
Wb_Par C
Wb_PtAy C
Wb_PtNy C
Wb_TxtCn C
Wb_TxtCnCnt C
Wb_TxtCnStr C
Wb_WcAy_OLE C
Wb_WcNy C
Wb_WcSy_OLE C
Wb_Ws C, A
Wb_WsNy C
Wb_Ws_BY_CD_NM C, D
Wb_Ws_Cd C, D
Wb_Ws_CdNy C
Wb_XAdd_Dbt C, E, D, F
Wb_XAdd_Dbtt C, E, A, F
Wb_XAdd_Dt C, G
Wb_XAdd_Wc C, D, D
Wb_XAdd_Ws C, D, F, F, D, D
Wb_XHas_OupNy_XAss C, H
Wb_XCls_NoSav C
Wb_XDlt_Wc C
Wb_XDlt_Ws C, A
Wb_XMax C
Wb_XNew_A1 C, D
Wb_XQuit C
Wb_XSav C
Wb_XSavAs C, A
Wb_XSet_WcTxtCn C, D
Wb_XVis C
Wb_XHas_Ws C, A
XX = CurWb()
End Sub

Private Sub Z()
Z_Wb_TxtCnCnt
Z_Wb_XSet_WcTxtCn
End Sub
