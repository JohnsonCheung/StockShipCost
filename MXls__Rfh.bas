Attribute VB_Name = "MXls__Rfh"
Option Compare Binary
Option Explicit
Sub Wc_XSet_FbCnStr(A As WorkbookConnection, Fb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
Dim Cn$
Const Ver$ = "0.0.1"
Select Case Ver
Case "0.0.1"
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, CStr(Fb), "Data Source=", ";")
Case "0.0.2"
    Cn = Fb_CnStr_OleAdo(Fb)
End Select
A.OLEDBConnection.Connection = Cn
End Sub

Function Fx_XRfh(A, Fb$) As Workbook
Set Fx_XRfh = Wb_XRfh(Fx_Wb(A), Fb)
End Function

Sub Ws_XRfh(A As Worksheet)
Itr_XDo A.QueryTables, "Qt_XRfh"
Itr_XDo A.PivotTables, "Pt_XRfh"
Itr_XDo A.ListObjects, "Lo_XRfh"
End Sub
Sub Lo_XRfh(A As ListObject)
A.Refresh
End Sub
Sub Qt_XRfh(A As excel.QueryTable)
A.BackgroundQuery = False
A.Refresh
End Sub
Function Wb_XCls_AllWc(A As Workbook)
'It does not work due to the ADOConnection is already closed, but
'the WFb is still in used.
'However, deleting the Wc works!!
Dim C As WorkbookConnection
Dim B As OLEDBConnection
For Each C In A.Connections
    If Not IsNothing(C.OLEDBConnection) Then
        Set B = A.Connections(1).OLEDBConnection
        'Cn_XCls A.Connections(1).OLEDBConnection.ADOConnection
        Cn_XCls C.OLEDBConnection.ADOConnection
    End If
Next
End Function
Function Wb_XRfh(A As Workbook, Optional Fb0 = "") As Workbook
Dim Fb$
Fb = DftStr(Fb0, CurDb.Name)
If A.Connections.Count = 0 Then FbWb_XRpl_Lo Fb, A
Wb_XRfh_FbCn A, Fb
Wb_XRfh_PivCaches A
Wb_XRfh_Sheets A
'Wb_Fmt_AllLo A
Set Wb_XRfh = A
If False Then
    'Not work
    Wb_XCls_AllWc A
Else
    'Ok
    Itr_XDo A.Connections, "Wc_XDlt"
End If
End Function
Sub Wb_XRfh_FbCn(A As Workbook, Fb$)
Dim C As WorkbookConnection
For Each C In A.Connections
    Wc_XSet_FbCnStr C, Fb
Next
End Sub
Function Wb_XRfh_FbCnStr(A As Workbook, Fb$) As Workbook
Itr_XDo_XP A.Connections, "Wc_XRfh_FbCnStr", Fb_CnStr_OleAdo(Fb)
Set Wb_XRfh_FbCnStr = A
End Function
Sub Wb_XRfh_Sheets(A As Workbook)
Dim W As Worksheet
For Each W In A.Sheets
    Ws_XRfh W
Next
End Sub
Sub Wb_XRfh_PivCaches(A As Workbook)
Dim C As PivotCache
For Each C In A.PivotCaches
    C.MissingItemsLimit = xlMissingItemsNone
    C.Refresh
Next
End Sub
Sub Pt_XRfh(A As excel.PivotTable)
A.Update
End Sub

