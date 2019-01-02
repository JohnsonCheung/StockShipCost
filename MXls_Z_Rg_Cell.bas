Attribute VB_Name = "MXls_Z_Rg_Cell"
Option Compare Binary
Option Explicit

Function Cell_AyH_WithValue(A As Range) As Variant()
If IsEmpty(RgA1(A).Value) Then Exit Function
Dim R&
For R = 2 To MaxRow
    If IsEmpty(RgRC(A, R, 1).Value) Then
        Cell_AyH_WithValue = Sq_Col(RgCRR(A, 1, 1, R - 1).Value, 1)
        Exit Function
    End If
Next
Stop
End Function

Sub Cell_XClr_Down(A As Range)
Cell_VBar_WithValue(A, AtLeastOneCell:=True).Clear
End Sub

Sub Cell_XFill_SeqDown(A As Range, FmNum&, ToNum&)
Ay_XPut_AtV SeqOfLng(FmNum, ToNum), A
End Sub

Function Cell_IsInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Function
R2 = R1 + Rg.Rows.Count
If R > R2 Then Exit Function
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Function
C2 = C1 + Rg.Columns.Count
If C > C2 Then Exit Function
Cell_IsInRg = True
End Function

Function Cell_IsInRgAp(A As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
Cell_IsInRgAp = Cell_IsInRgAv(A, Av)
End Function

Function Cell_IsInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If Cell_IsInRg(A, CvRg(V)) Then Cell_IsInRgAv = True: Exit Function
Next
End Function

Sub Cell_XLnk_Ws(A As Range, WsNy$())
Dim WsNm: WsNm = A.Value
If Not IsStr(WsNm) Then Exit Sub
If Not Ay_XHas(WsNy, WsNm) Then Exit Sub
With A.Hyperlinks
    If .Count > 0 Then .Delete
    .Add A, "", QQ_Fmt("'?'!A1", WsNm)
End With
End Sub

Sub Cell_XMge_Above(A As Range)
If Not IsEmpty(A.Value) Then Exit Sub
If A.MergeCells Then Exit Sub
If A.Row = 1 Then Exit Sub
If RgRC(A, 0, 1).MergeCells Then Exit Sub
RgMge RgCRR(A, 1, 0, 1)
End Sub

Function Cell_XReSz(A As Range, Sq) As Range
Set Cell_XReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function

Function Cell_VBar_WithValue(Cell As Range, Optional AtLeastOneCell As Boolean) As Range
If IsEmpty(Cell.Value) Then
    If AtLeastOneCell Then
        Set Cell_VBar_WithValue = RgA1(Cell)
    End If
    Exit Function
End If
Dim R1&: R1 = Cell.Row
Dim R2&
    If IsEmpty(RgRC(Cell, 2, 1)) Then
        R2 = Cell.Row
    Else
        R2 = Cell.End(xlDown).Row
    End If
Dim C%: C = Cell.Column
Set Cell_VBar_WithValue = Ws_CRR(Rg_Ws(Cell), C, R1, R2)
End Function
