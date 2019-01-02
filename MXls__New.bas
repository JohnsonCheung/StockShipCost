Attribute VB_Name = "MXls__New"
Option Compare Binary
Option Explicit
Function New_A1(Optional WsNm$ = "Sheet1") As Range
Set New_A1 = Ws_A1(New_Ws(WsNm))
End Function

Function New_Wb(Optional WsNm$) As Workbook
Dim O As Workbook
Dim X As excel.Application
Set O = New_Xls.Workbooks.Add
Set New_Wb = Ws_Wb(Ws_XSet_Nm(Wb_FstWs(O), WsNm))
End Function

Function New_Ws(Optional WsNm$) As Worksheet
Set New_Ws = Ws_XSet_Nm(Wb_FstWs(New_Wb), WsNm)
End Function

Property Get New_Xls() As excel.Application
Set New_Xls = CreateObject("Excel.Application") ' Don't use New Excel.Application
End Property
