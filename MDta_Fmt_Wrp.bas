Attribute VB_Name = "MDta_Fmt_Wrp"
Option Compare Binary
Option Explicit
Function WrpDrNRow%(WrpDr())
Dim Col, R%, M%
For Each Col In AyNz(WrpDr)
    M = Sz(Col)
    If M > R Then R = M
Next
WrpDrNRow = R
End Function

Function WrpDrPad(WrpDr, W%()) As Variant() _
'Some Cell in WrpDr may be an array, wraping each element to cell if their width can fit its W%(?)
Dim J%, Cell, O()
O = WrpDr
For Each Cell In AyNz(O)
    If IsArray(Cell) Then
        O(J) = AyWrpPad(Cell, W(J))
    End If
    J = J + 1
Next
WrpDrPad = O
End Function

Function WrpDrSq(WrpDr()) As Variant()
Dim O(), R%, C%, NCol%, NRow%, Cell, Col, NColi%
NCol = Sz(WrpDr)
NRow = WrpDrNRow(WrpDr)
ReDim O(1 To NRow, 1 To NCol)
C = 0
For Each Col In WrpDr
    C = C + 1
    If IsArray(Col) Then
        NColi = Sz(Col)
        For R = 1 To NRow
            If R <= NColi Then
                O(R, C) = Col(R - 1)
            Else
                O(R, C) = ""
            End If
        Next
    Else
        O(1, C) = Col
        For R = 2 To NRow
            O(R, C) = ""
        Next
    End If
Next
WrpDrSq = O
End Function

Function WrpDryWdt(WrpDry(), WrpWdt%) As Integer() _
'WrpDry is dry having 1 or more wrpCol, which mean need wrapping.
'WrpWdt is for wrpCol _
'WrpCol is col with each cell being array
'if maxWdt of array-ele of wrpCol has wdt > WrpWdt, use that wdt
'otherwise use WrpWdt
If Sz(WrpDry) = 0 Then Exit Function
Dim J%, Col()
For J = 0 To Dry_NCol(WrpDry) - 1
    Col = DryCol(WrpDry, J)
    If IsArray(Col(0)) Then
        Push WrpDryWdt, Ay_Wdt(AyFlat(Col))
    Else
        Push WrpDryWdt, Ay_Wdt(Col)
    End If
Next
End Function


Function Dr_FmtCellWrp(A, ColWdt%()) As String()
Dim X
For Each X In AyNz(A)
    Stop '
'    PushIAy Dr_FmtCellWrp, FmtssWrp(X, ColWdt)
Next
End Function

Function Dr_FmtWrp(WrpDr, W%()) As String() _
'Each Itm of WrpDr may be an array.  So a Dr_Fmt return Ly not string.
Dim Dr(): Dr = WrpDrPad(WrpDr, W)
Dim Sq(): Sq = WrpDrSq(Dr)
Dim Sq1(): Sq1 = SqAlign(Sq, W)
Dim Ly$(): Ly = Sq_Ly(Sq1)
PushIAy Dr_FmtWrp, Ly
End Function


Function Dry_FmtssWrp(A(), Optional WrpWdt% = 40) As String() _
'WrpWdt is for wrp-col.  If maxWdt of an ele of wrp-col > WrpWdt, use the maxWdt
Dim W%(), Dr, A1(), M$()
W = WrpDryWdt(A, WrpWdt)
For Each Dr In AyNz(A)
    M = Dr_FmtWrp(Dr, W)
    PushIAy Dry_FmtssWrp, M
Next
End Function

