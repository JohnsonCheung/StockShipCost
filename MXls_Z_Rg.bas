Attribute VB_Name = "MXls_Z_Rg"
Option Compare Binary
Option Explicit

Function CvRg(A) As Range
Set CvRg = A
End Function

Function RgA1(A As Range) As Range
Set RgA1 = RgRC(A, 1, 1)
End Function

Function Rg_A1LasCell(A As Range) As Range
Dim L As Range, R, C
Set L = A.SpecialCells(xlCellTypeLastCell)
R = L.Row
C = L.Column
Set Rg_A1LasCell = Ws_RCRC(Rg_Ws(A), A.Row, A.Column, R, C)
End Function

Function Rg_Adr$(A As Range)
Rg_Adr = "'" & Rg_Ws(A).Name & "'!" & A.Address
End Function

Sub Rg_RCRC_XAsg(A As Range, OR1, OC1, OR2, OC2)
OR1 = A.Row
OR2 = OR1 + A.Rows.Count - 1
OC1 = A.Column
OC2 = OC1 + A.Columns.Count - 1
End Sub

Sub Rg_XBdr_(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub

Sub Rg_XBdr_Around(A As Range)
Rg_XBdr_L A
Rg_XBdr_R A
Rg_XBdr_Top A
Rg_XBdr_Bottom A
End Sub

Sub Rg_XBdr_B(A As Range)
Rg_XBdr_L A
Rg_XBdr_R A
End Sub

Sub Rg_XBdr_Bottom(A As Range)
Rg_XBdr_ A, xlEdgeBottom
Rg_XBdr_ A, xlEdgeTop
End Sub

Sub Rg_XBdr_Inner(A As Range)
Rg_XBdr_ A, xlInsideHorizontal
Rg_XBdr_ A, xlInsideVertical
End Sub

Sub Rg_XBdr_Inside(A As Range)
Rg_XBdr_Inner A
End Sub

Sub Rg_XBdr_L(A As Range)
Rg_XBdr_Left A
End Sub

Sub Rg_XBdr_Left(A As Range)
Rg_XBdr_ A, xlEdgeLeft
If A.Column > 1 Then
    Rg_XBdr_ RgC(A, 0), xlEdgeRight
End If
End Sub

Sub Rg_XBdr_R(A As Range)
Rg_XBdr_Right A
End Sub

Sub Rg_XBdr_Right(A As Range)
Rg_XBdr_ A, xlEdgeRight
If A.Column < MaxCol Then
    Rg_XBdr_ RgC(A, A.Column + 1), xlEdgeLeft
End If
End Sub

Sub Rg_XBdr_Top(A As Range)
Rg_XBdr_ A, xlEdgeTop
If A.Row > 1 Then
    Rg_XBdr_ RgC(A, A.Row + 1), xlEdgeBottom
End If
End Sub

Function RgC(A As Range, C) As Range
Set RgC = RgCC(A, C, C)
End Function

Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, RgNRow(A), C2)
End Function

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function RgDecBtmR(A As Range, Optional By% = 1) As Range
Set RgDecBtmR = RgRR(A, 1, A.Rows.Count + 1)
End Function

Function RgEntC(A As Range, C) As Range
Set RgEntC = RgC(A, C).EntireColumn
End Function

Function RgEntRR(A As Range, R1, R2) As Range
Set RgEntRR = RgRR(A, R1, R2).EntireRow
End Function

Sub RgFillCol(A As Range)
Dim Rg As Range
Dim Sq()
Sq = N_SqV(A.Rows.Count)
CellSq_XReSz(A, Sq).Value = Sq
End Sub

Sub RgFillRow(A As Range)
Dim Rg As Range
Dim Sq()
Sq = N_SqH(A.Rows.Count)
CellSq_XReSz(A, Sq).Value = Sq
End Sub

Function RgFstC(A As Range) As Range
Set RgFstC = RgC(A, 1)
End Function

Function RgFstR(A As Range) As Range
Set RgFstR = RgR(A, 1)
End Function

Function RgIncTopR(A As Range, Optional By% = 1) As Range
Set RgIncTopR = RgRR(A, 1 - By, A.Rows.Count)
End Function

Function RgIsHBar(A As Range) As Boolean
RgIsHBar = A.Rows.Count = 1
End Function

Function RgIsVBar(A As Range) As Boolean
RgIsVBar = A.Columns.Count = 1
End Function

Function RgLasCol%(A As Range)
RgLasCol = A.Column + A.Columns.Count - 1
End Function

Function RgLasHBar(A As Range) As Range
Set RgLasHBar = RgR(A, RgNRow(A))
End Function

Function RgLasRow&(A As Range)
RgLasRow = A.Row + A.Rows.Count - 1
End Function

Function RgLasVBar(A As Range) As Range
Set RgLasVBar = RgC(A, RgNCol(A))
End Function

Sub Rg_XLnk_Ws(A As Range)
Dim R As Range
Dim WsNy$(): WsNy = Wb_WsNy(RgWb(A))
For Each R In A
    Cell_XLnk_Ws R, WsNy
Next
End Sub

Function Rg_Lo(A As Range, Optional LoNm$) As ListObject
Dim Ws As Worksheet: Set Ws = Rg_Ws(A)
Dim O As ListObject: Set O = Ws.ListObjects.Add(xlSrcRange, A, , xlYes)
Lo_XSet_Nm O, LoNm
Rg_XBdr_Around A
Set Rg_Lo = O
End Function

Sub RgMge(A As Range)
A.MergeCells = True
A.HorizontalAlignment = XlHAlign.xlHAlignCenter
A.VerticalAlignment = XlVAlign.xlVAlignCenter
End Sub

Function RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Function

Function RgNMoreBelow(A As Range, Optional N% = 1)
Set RgNMoreBelow = RgRR(A, 1, A.Rows.Count + N)
End Function

Function RgNMoreTop(A As Range, Optional N% = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, A.Rows.Count)
Set RgNMoreTop = O
End Function

Function RgNRow&(A As Range)
RgNRow = A.Rows.Count
End Function

Function RgR(A As Range, R)
Set RgR = RgRR(A, R, R)
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = Rg_Ws(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, RgNCol(A))
End Function

Function CellSq_XReSz(A As Range, Sq) As Range
Set CellSq_XReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function

Function RgSq(A As Range)
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        RgSq = O
        Exit Function
    End If
End If
RgSq = A.Value
End Function

Function RgVis(A As Range) As Range
Xls_XVis A.Application
Set RgVis = A
End Function

Function RgWb(A As Range) As Workbook
Set RgWb = Ws_Wb(Rg_Ws(A))
End Function

Function Rg_Ws(A As Range) As Worksheet
Set Rg_Ws = A.Parent
End Function

Sub RgeMgeV(A As Range)
Stop '?
End Sub


Private Sub Z_RgNMoreBelow()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = New_Ws
Set R = Ws.Range("A3:B5")
Set Act = RgNMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub

Private Sub Z()
Z_RgNMoreBelow
MXls_Z_Rg:
End Sub
