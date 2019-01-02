Attribute VB_Name = "MXls_Z_Ws"
Option Compare Binary
Option Explicit

Function Ws_C(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set Ws_C = R.EntireColumn
End Function

Function Ws_CC(A As Worksheet, C1, C2) As Range
Set Ws_CC = Ws_RCC(A, 1, C1, C2).EntireColumn
End Function
Function Ws_A1(A As Worksheet) As Range
Set Ws_A1 = A.Cells(1, 1)
End Function


Sub Ws_ClrLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = Itr_Into(A.ListObjects, Ay)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub

Sub Ws_ClsNoSav(A As Worksheet)
Wb_XCls_NoSav Ws_Wb(A)
End Sub
Property Get CurWs() As Worksheet
Set CurWs = CurXls.ActiveSheet
End Property


Function Ws_CRR(A As Worksheet, C, R1, R2) As Range
Set Ws_CRR = Ws_RCRC(A, R1, C, R2, C)
End Function

Function Ws_DftLoNm$(A As Worksheet, Optional LoNm0$)
Dim LoNm$: LoNm = DftStr(LoNm0, "Table")
Dim J%
For J = 1 To 999
    If Not Ws_XHas_LoNm(A, LoNm) Then Ws_DftLoNm = LoNm: Exit Function
    LoNm = Nm_NxtSeqNm(LoNm)
Next
Stop
End Function

Function WsDlt(A As Workbook, WsIx) As Boolean
If Wb_XHas_Ws(A, WsIx) Then Wb_Ws(A, WsIx).Delete: Exit Function
WsDlt = True
End Function

Function WsDtaRg(A As Worksheet) As Range
Dim R, C
With Ws_LasCell(A)
   R = .Row
   C = .Column
End With
If R = 1 And C = 1 Then Exit Function
Set WsDtaRg = Ws_RCRC(A, 1, 1, R, C)
End Function

Function Ws_FstLo(A As Worksheet) As ListObject
Set Ws_FstLo = Itr_FstItm(A.ListObjects)
End Function

Function Ws_XHas_LoNm(A As Worksheet, LoNm$) As Boolean
Ws_XHas_LoNm = Itr_XHas_Nm(A.ListObjects, LoNm)
End Function

Function Ws_LasCell(A As Worksheet) As Range
Set Ws_LasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function Ws_LasCno%(A As Worksheet)
Ws_LasCno = Ws_LasCell(A).Column
End Function

Function Ws_LasRno%(A As Worksheet)
Ws_LasRno = Ws_LasCell(A).Row
End Function

Function Ws_Lo(A As Worksheet, LoNm$) As ListObject
Set Ws_Lo = A.ListObjects(LoNm)
End Function

Function Ws_PtAy(A As Worksheet) As PivotTable()
Dim O() As PivotTable, Pt As PivotTable
For Each Pt In A.PivotTables
    PushObj O, Pt
Next
Ws_PtAy = O
End Function

Function WsPtNy(A As Worksheet) As String()
Dim Pt As PivotTable
For Each Pt In A.PivotTables
    PushI WsPtNy, Pt.Name
Next
End Function

Function Ws_RC(A As Worksheet, R, C) As Range
Set Ws_RC = A.Cells(R, C)
End Function

Function Ws_RCC(A As Worksheet, R, C1, C2) As Range
Set Ws_RCC = Ws_RCRC(A, R, C1, R, C2)
End Function

Function Ws_RCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set Ws_RCRC = A.Range(Ws_RC(A, R1, C1), Ws_RC(A, R2, C2))
End Function

Function Ws_RR(A As Worksheet, R1, R2) As Range
Set Ws_RR = A.Range(Ws_RC(A, R1, 1), Ws_RC(A, R2, 1)).EntireRow
End Function

Function Ws_XSet_LoNm(A As Worksheet, Nm$) As Worksheet
Dim Lo As ListObject
Set Lo = CvNothing(Itr_FstItm(A.ListObjects))
If Not IsNothing(Lo) Then Lo.Name = "T_" & Nm
Set Ws_XSet_LoNm = A
End Function

Function Ws_XSet_Nm(A As Worksheet, Nm) As Worksheet
If Nm <> "" Then
    If Not Wb_XHas_Ws(Ws_Wb(A), Nm) Then A.Name = Nm
End If
Set Ws_XSet_Nm = A
End Function

Function Ws_Sq(A As Worksheet) As Variant()
Ws_Sq = WsDtaRg(A).Value
End Function

Function Ws_XVis(A As Worksheet) As Worksheet
Xls_XVis A.Application
Set Ws_XVis = A
End Function


Function Ws_Wb(A As Worksheet) As Workbook
Set Ws_Wb = A.Parent
End Function
