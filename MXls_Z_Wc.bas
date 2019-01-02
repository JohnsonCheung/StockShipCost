Attribute VB_Name = "MXls_Z_Wc"
Option Compare Binary
Option Explicit

Function Wc_Ws(A As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet, Lo As ListObject, Qt As QueryTable
Set Wb = A.Parent
Set Ws = Wb_XAdd_Ws(Wb, A.Name)
Ws.Name = A.Name
Wc_XPut_At A, Ws_A1(Ws)
Set Wc_Ws = Ws
End Function

Sub Wc_XDlt(A As WorkbookConnection)
A.Delete
End Sub

Sub Wc_XPut_At(A As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = Rg_Ws(At).ListObjects.Add(SourceType:=0, Source:=A.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = A.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = TblNm_LoNm(A.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub Wc_XRfh_Fb(A As WorkbookConnection, Fb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
Wc_XSet_FbCnStr A, Fb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub

Private Sub ZZ()
Dim A As WorkbookConnection
Dim B As Range
Dim C
Dim XX
Wc_XDlt A
Wc_XPut_At A, B
Wc_XRfh_Fb A, C
End Sub

Private Sub Z()
End Sub
