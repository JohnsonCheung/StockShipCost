Attribute VB_Name = "MXls__Dta"
Option Compare Binary
Option Explicit

Function Drs_Lo(A As Drs, At As Range) As ListObject
Set Drs_Lo = Rg_Lo(Sq_XPut_AtCell(Drs_Sq(A), At))
End Function

Function Drs_Lo_FMT(A As Drs, At As Range, Lo_Fmtr$()) As ListObject
Set Drs_Lo_FMT = Lo_Fmt(Drs_Lo(A, At), Lo_Fmtr)
End Function

Function Drs_Rg(A As Drs, At As Range) As Range
Set Drs_Rg = Sq_XPut_AtCell(Drs_Sq(A), At)
End Function

Function Drs_Ws1(A As Drs) As Worksheet
Set Drs_Ws1 = Sq_Ws(Drs_Sq(A))
End Function

Function Drs_Ws(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = New_Ws(WsNm)
Drs_Lo A, Ws_A1(O)
Set Drs_Ws = O
End Function

Function Drs_Ws_Vis(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Set Drs_Ws_Vis = Ws_XVis(Drs_Ws(A, WsNm))
End Function

Function Drs_XBrw_Ws(A As Drs) As Worksheet
Set Drs_XBrw_Ws = Drs_Ws_Vis(A)
End Function

Function Drs_XPut_At(A As Drs, At As Range) As Range
Set Drs_XPut_At = Sq_XPut_AtCell(Drs_Sq(A), At)
End Function

Sub Drs_XPut_Ws(A As Drs, PutWs As Worksheet)
Drs_XPut_At A, Ws_A1(PutWs)
End Sub

Function Dry_Ws(Dry, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = New_Ws(WsNm)
Dry_XPut_Rg Dry, Ws_A1(O)
Set Dry_Ws = O
End Function

Function Dry_XPut_Rg(A, At As Range) As Range
Set Dry_XPut_Rg = Sq_XPut_AtCell(Dry_Sq(A), At)
End Function

Function Ds_Ws(A As Ds) As Worksheet
Dim O As Worksheet: Set O = New_Ws
Ws_A1(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = Ws_RC(O, 2, 1)
Dim I
For Each I In AyNz(A.DtAy)
    Dt_XPut_AtCell CvDt(I), At
Next
Set Ds_Ws = O
End Function

Function Ds_Wb(A As Ds) As Workbook
Dim O As Workbook
Set O = New_Wb
With Wb_FstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Dim I
For Each I In AyNz(A.DtAy)
    Wb_XAdd_Dt O, CvDt(I)
Next
Set Ds_Wb = O
End Function

Function DtAp_Wb(ParamArray DtAp()) As Workbook
Dim DtAy() As Dt, Av()
Av = DtAp
DtAy = Oy_Into(Av, DtAy)
End Function

Function Dt_XPut_AtCell1(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = Drs_Fmt(Dt_Drs(A))
Ay_XPut_AtV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set Dt_XPut_AtCell1 = At
End Function
Function Cell_XNxt_RowCol(A As Range) As Range

End Function

Sub Dt_XPut_AtCell(A As Dt, AtCell As Range)
Rg_Lo _
    Sq_XPut_AtCell( _
    Drs_Sq( _
    Dt_Drs(A)), AtCell)
End Sub
Function Cell_XAdd_NRow(A As Range, NRow) As Range
Stop '
End Function

Function Dt_XAdd_Wb(A As Dt, Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = Wb_XAdd_Ws(Wb, A.DtNm)
Drs_Lo Dt_Drs(A), Ws_A1(O)
Set Dt_XAdd_Wb = O
End Function

Function Dt_Ws(A As Dt) As Worksheet
Set Dt_Ws = Dt_XPut_Ws(A, New_Ws)
End Function

Function Dt_XPut_Ws(A As Dt, PutWs As Worksheet) As Worksheet
Set Dt_XPut_Ws = PutWs
Dt_XPut_AtCell A, Ws_A1(PutWs)
Ws_XSet_Nm PutWs, A.DtNm
End Function

Function Sq_XPut_AtCell(A, AtCell As Range) As Range
Dim O As Range
Set O = CellSq_XReSz(AtCell, A)
O.MergeCells = False
O.Value = A
Set Sq_XPut_AtCell = O
End Function

Private Sub ZZ_Ds_Ws()
Ws_XVis Ds_Ws(Samp_Ds)
End Sub

Private Sub ZZ_Ds_Wb()
Dim Wb As Workbook
Stop '
'Set Wb = Ds_Wb(Db_Ds(CurDb, "Permit PermitD"))
Wb_XVis Wb
Stop
Wb.Close False
End Sub

Private Sub Z_Ds_Wb()
Dim Wb As Workbook
Set Wb = Ds_Wb(Db_Ds(CurDb, "Permit PermitD"))
Wb_XVis Wb
Stop
Wb.Close False
End Sub

Private Sub ZZ()
Dim A As Drs
Dim B As Range
Dim C()
Dim D$()
Dim E As Dt
Dim F As Worksheet
Dim H
Dim I As Ds
Dim J%
Dim K As Workbook
Dim XX
Drs_Lo A, B
Drs_Lo_FMT A, B, D
Drs_Rg A, B
End Sub

Private Sub Z()
Z_Ds_Wb
End Sub
