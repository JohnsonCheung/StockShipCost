Attribute VB_Name = "MXls_Z_Lo"
Option Compare Binary
Option Explicit
Const CMod$ = "MXls_Z_Lo."

Function CvLo(A) As ListObject
Set CvLo = A
End Function

Function Lo_AllCol(A As ListObject) As Range
Set Lo_AllCol = LoCC(A, 1, Lo_NCol(A))
End Function

Function Lo_AllEntCol(A As ListObject) As Range
Set Lo_AllEntCol = Lo_AllCol(A).EntireColumn
End Function

Sub Lo_XAutoFit(A As ListObject)
Dim C As Range: Set C = Lo_AllEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = RgEntC(C, J)
   If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Sub Lo_XBdr_Around(A As ListObject)
Dim R As Range
Set R = RgNMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgNMoreBelow(R)
Rg_XBdr_Around R
End Sub

Sub Lo_XBrw(A As ListObject)
Drs_XBrw Lo_Drs(A)
End Sub

Function LoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = Lo_R1(A, InclHdr)
R2 = Lo_R2(A, InclTot)
mC1 = LoC_WsCno(A, C1)
mC2 = LoC_WsCno(A, C2)
Set LoCC = Ws_RCRC(Lo_Ws(A), R1, mC1, R2, mC2)
End Function


Function LoCol(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R As Range
Set R = A.ListColumns(C).DataBodyRange
If Not InclTot And Not InclHdr Then
    Set LoCol = R
    Exit Function
End If

If InclTot Then Set LoCol = RgDecBtmR(R, 1)
If InclHdr Then Set LoCol = RgIncTopR(R, 1)
End Function

Function LoCol_Sy(A As ListObject, C) As String()
LoCol_Sy = Sq_Col_Sy(A.ListColumns(C).DataBodyRange.Value, 1)
End Function

Sub LoCol_XLnk_Ws(A As ListObject, C)
Rg_XLnk_Ws LoCol(A, C)
End Sub

Function Ws_XCrt_Lo(A As Worksheet, Optional LoNm$) As ListObject
Dim R As Range: Set R = WsDtaRg(A)
If IsNothing(R) Then Exit Function
Dim O As ListObject: Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), , xlYes)
If LoNm <> "" Then O.Name = LoNm
Lo_XAutoFit O
Set Ws_XCrt_Lo = O
End Function

Sub Lo_XDlt(A As ListObject)
Dim R As Range, R1, C1, R2, C2, Ws As Worksheet
Set Ws = Lo_Ws(A)
Set R = RgNMoreBelow(RgNMoreTop(A.DataBodyRange))
Rg_RCRC_XAsg R, R1, C1, R2, C2
A.QueryTable.Delete
Ws_RCRC(Ws, R1, C1, R2, C2).ClearContents
End Sub

Function Lo_Drs(A As ListObject) As Drs
Set Lo_Drs = New_Drs(Lo_Fny(A), Lo_Dry(A))
End Function

Function Lo_Dry(A As ListObject) As Variant()
Lo_Dry = Sq_Dry(Lo_Sq(A))
End Function

Function LoFF_Dry(A As ListObject, FF) As Variant() _
' Return as many column as fields in [FF] from Lo[A]
Dim IxAy&(), Dry(): GoSub X_IxAy_Dry
Dim Dr
For Each Dr In AyNz(Dry)
    PushI LoFF_Dry, DrSel(Dr, IxAy)
Next
Exit Function
X_IxAy_Dry:
    Dim Fny$()
    Fny = Lo_Fny(A)
    Dry = Lo_Dry(A)
    IxAy = Ay_IxAy(Fny, Ssl_Sy(FF))
    Return
End Function

Function Lo_DtaAdr$(A As ListObject)
Lo_DtaAdr = Rg_Adr(A.DataBodyRange)
End Function

Sub Lo_XEns_NRow(A As ListObject, NRow&)
Lo_XMin A
Exit Sub
If NRow > 1 Then
    Debug.Print A.InsertRowRange.Address
    Stop
End If
End Sub

Function Lo_EntCol(A As ListObject, C) As Range
Set Lo_EntCol = LoCol(A, C).EntireColumn
End Function

Function Lo_FbtStr$(A As ListObject)
Lo_FbtStr = Qt_FbtStr(A.QueryTable)
End Function

Function Lo_Fny(A As ListObject) As String()
Lo_Fny = Itr_Ny(A.ListColumns)
End Function

Function Lo_XHas_Col(A As ListObject, C) As Boolean
Lo_XHas_Col = Itr_XHas_Nm(A.ListColumns, C)
End Function

Function Lo_XHas_Fny(A As ListObject, Fny$()) As Boolean
Lo_XHas_Fny = Ay_XHasAy(Lo_Fny(A), Fny)
End Function

Function Lo_XHas_NoDta(A As ListObject) As Boolean
Lo_XHas_NoDta = IsNothing(A.DataBodyRange)
End Function

Function LoC_HdrCell(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set LoC_HdrCell = RgRC(Rg, 1, 1)
End Function

Sub Lo_XKeep_FstCol(A As ListObject)
Dim J%
For J = A.ListColumns.Count To 2 Step -1
    A.ListColumns(J).Delete
Next
End Sub

Sub Lo_XKeep_FstRow(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub

Function Lo_NCol%(A As ListObject)
Lo_NCol = A.ListColumns.Count
End Function

Function LoNm_TblNm$(A)
If Not XHas_Pfx(A, "T_") Then Stop
LoNm_TblNm = "@" & Mid(A, 3)
End Function

Function Lo_Pc(A As ListObject) As PivotCache
Dim O As PivotCache
Set O = Lo_Wb(A).PivotCaches.Create(xlDatabase, A.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set Lo_Pc = O
End Function

Function Lo_Pt(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
If Lo_Wb(A).FullName <> RgWb(At).FullName Then Stop: Exit Function
Dim O As PivotTable
Set O = Lo_Pc(A).CreatePivotTable(TableDestination:=At, TableName:=Lo_PtNm(A))
With O
    .ShowDrillIndicators = False
    .InGridDropZones = False
    .RowAxisLayout xlTabularRow
End With
O.NullString = ""
PtFFSetOri O, Rowss, xlRowField
PtFFSetOri O, Colss, xlColumnField
PtFFSetOri O, Pagss, xlPageField
PtFFSetOri O, Dtass, xlDataField
Set Lo_Pt = O
End Function

Function Lo_PtNm$(A As ListObject)
If Left(A.Name, 2) <> "T_" Then Stop
Dim O$: O = "P_" & Mid(A.Name, 3)
Lo_PtNm = AyNxtNm(Wb_PtNy(Lo_Wb(A)), O)
End Function

Function Lo_Qt(A As ListObject) As QueryTable
On Error Resume Next
Set Lo_Qt = A.QueryTable
End Function

Function Lo_R1&(A As ListObject, Optional InclHdr As Boolean)
If Lo_XHas_NoDta(A) Then
   Lo_R1 = A.ListColumns(1).Range.Row + 1
   Exit Function
End If
Lo_R1 = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function Lo_R2&(A As ListObject, Optional InclTot As Boolean)
If Lo_XHas_NoDta(A) Then
   Lo_R2 = Lo_R1(A)
   Exit Function
End If
Lo_R2 = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Sub LoFb_XReset(A As ListObject, Fb$)
'When Lo_XRfh, if the fields of Db's table has been reorder, the Lo will not follow the order
'Delete the Lo, add back the Wc then Wc_XPut_At to reset the Lo
'Tbl_Nm : from Lo.Name = T_XXX is the key to get the table name.
'Fb    : use WFb
Dim LoNm$, T$, At As Range, Wb As Workbook
Set Wb = Lo_Wb(A)
Set At = RgRC(A.DataBodyRange, 0, 1)
LoNm = A.Name
T = LoNm_TblNm(LoNm)
Lo_XDlt A
Wc_XPut_At Wb_XAdd_Wc(Wb, WFb, T), At
End Sub

Function LoFF_Drs(A As ListObject, FF) As Drs
Dim Fny$(): Fny = Lin_TermAy(FF)
Set LoFF_Drs = New_Drs(Fny, Sq_Dry_SEL(Lo_Sq(A), Ay_IxAy(Lo_Fny(A), Fny)))
End Function

Function Lo_XSet_Nm(A As ListObject, LoNm$) As ListObject
If LoNm <> "" Then
    If Ws_XHas_LoNm(A, LoNm) Then XThw CSub, "New-LoNm already exist", "Ws Old-LoNm New-LoNm"
    A.Name = LoNm
End If
Set Lo_XSet_Nm = A
End Function

Function Lo_Sq(A As ListObject)
Lo_Sq = A.DataBodyRange.Value
End Function

Function Lo_XVis(A As ListObject) As ListObject
Xls_XVis A.Application
Set Lo_XVis = A
End Function

Function Lo_Wb(A As ListObject) As Workbook
Set Lo_Wb = Ws_Wb(Lo_Ws(A))
End Function

Function Lo_Ws(A As ListObject) As Worksheet
Set Lo_Ws = A.Parent
End Function

Function LoC_WsCno%(A As ListObject, Col)
LoC_WsCno = A.ListColumns(Col).Range.Column
End Function

Function TblNm_LoNm$(Tbl_Nm)
TblNm_LoNm = "T_" & XRmv_FstNonLetter(Tbl_Nm)
End Function

Private Sub ZZ_Lo_XKeep_FstCol()
Lo_XKeep_FstCol Lo_XVis(SampLo)
End Sub

Private Sub Z_Lo_XAutoFit()
Dim Ws As Worksheet: Set Ws = New_Ws
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
Lo_XAutoFit Ws_XCrt_Lo(Ws)
Ws_ClsNoSav Ws
End Sub

Private Sub Z_Lo_XBrw()
Dim O As ListObject: Set O = SampLo
Lo_XBrw O
Stop
End Sub

Private Sub Z_Lo_Pt()
Dim At As Range, Lo As ListObject
Set Lo = SampLo
Set At = RgVis(Ws_A1(Wb_XAdd_Ws(Lo_Wb(Lo))))
PtVis Lo_Pt(Lo, At, "A B", "C D", "F", "E")
Stop
End Sub

Private Sub Z_LoFb_XReset()
Dim Wb As Workbook, LoAy() As ListObject
Set Wb = Fx_Wb("C:\users\user\desktop\a.xlsx")
Wb_XVis Wb
LoAy = Wb_LoAy(Wb)
'LoFb_XReset LoAy(0)
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As ListObject
Dim C As Boolean
Dim D As Worksheet
Dim E$
Dim F&
Dim G$()
Dim H As Range
CvLo A
Lo_AllCol B
Lo_AllEntCol B
Lo_XAutoFit B
Lo_XBdr_Around B
Lo_XBrw B
LoCC B, A, A, C, C
LoCol B, A, C, C
LoCol_Sy B, A
LoNm_TblNm A
Lo_Pc B
Lo_Pt B, H, E, E, E, E
Lo_PtNm B
Lo_Qt B
Lo_R1 B, C
Lo_R2 B, C
LoFb_XReset B, E
LoFF_Drs B, A
Lo_XSet_Nm B, E
Lo_Sq B
Lo_XVis B
Lo_Wb B
Lo_Ws B
LoC_WsCno B, A
TblNm_LoNm A
End Sub

Private Sub Z()
Z_Lo_XAutoFit
Z_Lo_XBrw
Z_Lo_Pt
Z_LoFb_XReset
End Sub
