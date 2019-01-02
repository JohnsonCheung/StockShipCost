Attribute VB_Name = "MXls__Sq"
Option Compare Binary
Option Explicit

Function New_Sq(R&, C&) As Variant()
Dim O()
ReDim O(1 To R, 1 To C)
New_Sq = O
End Function


Function Sq_Add_SngQuote(A)
Dim NC%, C%, R&, O
O = A
NC = UBound(A, 2)
For R = 1 To UBound(A, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
Sq_Add_SngQuote = O
End Function

Sub Sq_Brw(A)
Dry_XBrw Sq_Dry(A)
End Sub

Function Sq_Col(A, C%) As Variant()
Dim O()
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C)
Next
Sq_Col = O
End Function

Function Sq_Col_Into(A, C%, OInto) As String()
Dim O
O = OInto
Erase O
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C%)
Next
Sq_Col_Into = O
End Function

Function Sq_Col_Sy(A, C%) As String()
Sq_Col_Sy = Sq_Col_Into(A, C, EmpSy)
End Function
Function Sq_Col_Sy_Fst(A) As String()
Sq_Col_Sy_Fst = Sq_Col_Sy(A, 1)
End Function

Function SqR_Dr(A, R) As Variant()
Dim C%
For C = 1 To UBound(A, 2)
    PushI SqR_Dr, A(R, C)
Next
End Function


Function Sq_Dr_CNOAY(A, R&, CnoAy&()) As Variant()
Dim J%
ReDim mCnoAy(UBound(A, 2) - 1)
For J = 0 To UB(mCnoAy)
    mCnoAy(J) = J + 1
Next
Dim UCol%
   UCol = UB(mCnoAy)
Dim O()
   ReDim O(UCol)
   Dim C%
   For J = 0 To UCol
       C = mCnoAy(J)
       O(J) = A(R, C)
   Next
Sq_Dr_CNOAY = O
End Function

Function Sq_Dry(A) As Variant()
If Not IsArray(A) Then
    Sq_Dry = Array(Array(A))
    Exit Function
End If
Dim R&
For R = 1 To UBound(A, 1)
    PushI Sq_Dry, SqR_Dr(A, R)
Next
End Function

Function Sq_Ins_Dr(A, Dr, Optional Row& = 1)
Dim O(), C%, R&, NC%, NR&
NC = Sq_NCol(A)
NR = Sq_NRow(A)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = A(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = A(R, C)
    Next
Next
Sq_Ins_Dr = O
End Function

Function Sq_IsEmp(Sq) As Boolean
Sq_IsEmp = True
On Error GoTo X
Dim A
If UBound(Sq, 1) < 0 Then Exit Function
If UBound(Sq, 2) < 0 Then Exit Function
Sq_IsEmp = False
Exit Function
X:
End Function

Function Sq_IsEq(A, B) As Boolean
Dim NR&, NC&
NR = UBound(A, 1)
NC = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If NC <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To NC
        If A(R, C) <> B(R, C) Then
            Exit Function
        End If
    Next
Next
Sq_IsEq = True
End Function

Function Sq_Ly(A) As String()
Dim R%
For R = 1 To UBound(A, 1)
    Push Sq_Ly, JnSpc(SqR_Dr(A, R))
Next
End Function

Function Sq_NCol&(A)
On Error Resume Next
Sq_NCol = UBound(A, 2)
End Function

Function Sq_A1(A, Optional WsNm$ = "Data") As Range
Set Sq_A1 = Sq_XPut_AtCell(A, New_A1(WsNm))
End Function

Function Sq_Lo(A, Optional WsNm$ = "Data") As ListObject
Set Sq_Lo = Rg_Lo(Sq_A1(A, WsNm))
End Function

Function Sq_Ws(A, Optional WsNm$, Optional LoNm$) As Worksheet
Dim Lo As ListObject
Set Lo = Sq_Lo(A, New_A1(WsNm))
Lo_XSet_Nm Lo, LoNm
Set Sq_Ws = Lo_Ws(Lo)
End Function

Function Sq_NRow&(A)
On Error Resume Next
Sq_NRow = UBound(A, 1)
End Function

Function Sq_Dry_WH(A, R&, ColIxAy&()) As Variant()
Dim Ix
For Each Ix In ColIxAy
    Push Sq_Dry_WH, A(R, Ix + 1)
Next
End Function

Function Sq_Lo_RPL(A, Lo As ListObject) As ListObject
Dim LoNm$, At As Range
LoNm = Lo.Name
Set At = Lo.Range
Lo.Delete
Set Sq_Lo_RPL = Lo_XSet_Nm(Rg_Lo(Sq_XPut_AtCell(A, At)), LoNm)
End Function

Function Sq_Dry_SEL(A, ColIxAy&()) As Variant()
Dim R&
For R = 1 To Sq_NRow(A)
    PushI Sq_Dry_SEL, Sq_Dry_WH(A, R, ColIxAy)
Next
End Function

Sub Sq_Set_Row(OSq, R&, Dr)
Dim J%
For J = 0 To UB(Dr)
    OSq(R, J + 1) = Dr(J)
Next
End Sub

Function SqSyV(A) As String()
SqSyV = Sq_Col_Sy(A, 1)
End Function

Function Sq_Transpose(A) As Variant()
Dim NR&, NC&
NR = Sq_NRow(A): If NR = 0 Then Exit Function
NC = Sq_NCol(A): If NC = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To NC, 1 To NR)
For J = 1 To NR
    For I = 1 To NC
        O(I, J) = A(J, I)
    Next
Next
Sq_Transpose = O
End Function

Private Sub ZZ()
Dim A&
Dim B As Variant
Dim C$
Dim D%
Dim E&()
Dim F As ListObject
New_Sq A, A
Sq_A1 B, C
Sq_Add_SngQuote B
Sq_Brw B
Sq_Col B, D
Sq_Col_Into B, D, B
Sq_Col_Sy B, D
Sq_Dry B
Sq_Col_Sy_Fst B
Sq_Ins_Dr B, B, A
Sq_IsEmp B
Sq_IsEq B, B
Sq_Ly B
Sq_NCol B
Sq_A1 B, C
Sq_Lo B, C
Sq_Ws B
Sq_NRow B
Sq_Dry_WH B, A, E
Sq_Lo_RPL B, F
Sq_Dry_SEL B, E
Sq_Set_Row B, A, B
SqSyV B
Sq_Transpose B
Sq_Ws B, C
End Sub

Private Sub Z()
End Sub
