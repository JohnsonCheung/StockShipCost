Attribute VB_Name = "MXls__MinimizeLo"
Option Compare Binary
Option Explicit
Sub Lo_XMin(A As ListObject)
If XTak_FstTwoChr(A.Name) <> "T_" Then Exit Sub
Dim R1 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    RgRR(R1, 2, R1.Rows.Count).EntireRow.Delete
End If
End Sub

Private Sub Ws_XMin_AllLo(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
If XTak_FstTwoChr(A.CodeName) <> "Ws" Then Exit Sub
Dim L As ListObject
For Each L In A.ListObjects
    Lo_XMin L
Next
End Sub

Function Wb_XMin_AllLo(A As Workbook) As Workbook
Dim Ws As Worksheet
For Each Ws In A.Sheets
    Ws_XMin_AllLo Ws
Next
Set Wb_XMin_AllLo = A
End Function

Sub Fx_XMin_AllLo(A)
Wb_XCls_NoSav Wb_XSav(Wb_XMin_AllLo(Fx_Wb(A)))
End Sub
