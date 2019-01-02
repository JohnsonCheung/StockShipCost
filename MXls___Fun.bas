Attribute VB_Name = "MXls___Fun"
Attribute VB_Description = "aaa"
Option Compare Binary
Option Explicit

Sub Ay_XPut_Col(A, At As Range)
Dim Sq()
Sq = AySqV(A)
CellSq_XReSz(At, Sq).Value = Sq
End Sub

Sub Ay_XPut_LoCol(A, Lo As ListObject, ColNm$)
Dim At As Range, C As ListColumn, R As Range
'Ay_XDmp Lo_Fny(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
Ay_XPut_Col A, At
End Sub

Sub AyPutRow(A, At As Range)
Dim Sq()
Sq = AySqH(A)
CellSq_XReSz(At, Sq).Value = Sq
End Sub

Function Ay_XPut_AtH(A, At As Range) As Range
Set Ay_XPut_AtH = Sq_XPut_AtCell(AySqH(A), At)
End Function

Function Ay_XPut_AtV(A, At As Range) As Range
Set Ay_XPut_AtV = Sq_XPut_AtCell(AySqV(A), At)
End Function

Function Ay_Ws(A, Optional WsNm$) As Worksheet
Dim O As Worksheet: Set O = New_Ws(WsNm)
Sq_XPut_AtCell AySqV(A), Ws_A1(O)
Set Ay_Ws = O
End Function

Function AyAB_Sq_Hori(A, B, N1$, N2$)
AyAB_XSet_SamMax A, B
Dim N&, O(), J&
N = Sz(A)
ReDim O(1 To 2, 1 To N + 1)
For J = 0 To N - 1
    O(1, J + 2) = A(J)
    O(2, J + 2) = B(J)
Next
O(1, 1) = N1
O(2, 1) = N2
AyAB_Sq_Hori = O
End Function
Function AyAB_Sq_Vert(A, B, N1$, N2$)
AyAB_XSet_SamMax A, B
Dim N&, O(), J&
N = Sz(A)
ReDim O(1 To N + 1, 1 To 2)
For J = 0 To N - 1
    O(J + 2, 1) = A(J)
    O(J + 2, 2) = B(J)
Next
O(1, 1) = N1
O(1, 2) = N2
AyAB_Sq_Vert = O
End Function

Function AyAB_Ws_Hori(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional WsNm$ = "", Optional LoNm$ = "AyAB") As Worksheet
Set AyAB_Ws_Hori = Sq_Ws(AyAB_Sq_Hori(A, B, N1, N2), WsNm, LoNm)
End Function

Function AyAB_Ws_Vert(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional WsNm$ = "", Optional LoNm$ = "AyAB") As Worksheet
Set AyAB_Ws_Vert = Sq_Ws(AyAB_Sq_Vert(A, B, N1, N2), WsNm, LoNm)
End Function


Function DicWb(A As Dictionary) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass DicAllKeyIsNm(A)
Ass DicAllValIsStr(A)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook
Set O = New_Wb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        Set DicWb = O
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = Lines_SqV(A(K))
Next
X: Set DicWb = O
End Function

Function DicWs(A As Workbook, Optional InclDicValTy As Boolean) As Worksheet
Set DicWs = Drs_Ws(Dic_Drs(A, InclDicValTy))
End Function

Function DicWs_XVis(A As Dictionary) As Worksheet
Dim O As Worksheet
   Set O = DicWs(A)
   Ws_XVis O
Set DicWs_XVis = O
End Function

Function DtWs(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = New_Ws(A.DtNm)
Drs_Lo Dt_Drs(A), Ws_A1(O)
Set DtWs = O
If Vis Then Ws_XVis O
End Function

Function FmlNy(A$) As String()
FmlNy = MacroNy(A)
End Function

Sub Lc_XSet_TotLnk(A As ListColumn)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.DataBodyRange
Set Ws = Rg_Ws(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Function LyWs(Ly$(), Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = New_Ws()
Ay_XPut_AtV Ly, Ws_A1(O)
Set LyWs = O
End Function

Property Get MaxCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(CurXls.Version = "16.0", 16384, 255)
End If
MaxCol = C
End Property

Property Get MaxRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(CurXls.Version = "16.0", 1048576, 65535)
End If
MaxRow = R
End Property

Function N_SqH(N%) As Variant()
Dim O(), J%
ReDim O(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = J
Next
N_SqH = O
End Function

Function N_SqV(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = J
Next
N_SqV = O
End Function

Function N_ZerFill$(N, NDig%)
N_ZerFill = Format(N, String(NDig, "0"))
End Function

Private Function Pj_Ffn$(A As VBProject)
On Error Resume Next
Pj_Ffn = A.FileName
End Function

Function S1S2AyWs(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set S1S2AyWs = Sq_Ws(S1S2AySq(A, Nm1, Nm2))
End Function

Private Sub Z_AyAB_Ws_Hori()
GoTo ZZ
Dim A, B
ZZ:
    A = Ssl_Sy("A B C D E")
    B = Ssl_Sy("1 2 3 4 5")
    Ws_XVis AyAB_Ws_Hori(A, B)
End Sub

Private Sub Z_Fb_Wb_OUP_TBL()
GoTo ZZ
ZZ:
    Dim W As Workbook
    'Set W = Fb_Wb_OUP_TBL(WFb)
    'Wb_XVis W
    Stop
    'W.Close False
    Set W = Nothing
End Sub
