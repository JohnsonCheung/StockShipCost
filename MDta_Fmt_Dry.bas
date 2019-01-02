Attribute VB_Name = "MDta_Fmt_Dry"
Option Compare Binary
Option Explicit

Function Dry_Fmt(A, Optional MaxColWdt% = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean, Optional HidGrid As Boolean) As String()
If Sz(A) = 0 Then Exit Function
Dim A1(): A1 = Dry_StrCellDry(A, ShwZer) ' Convert each cell in Dry-A into string
Dim W%(): W = Dry_WdtAy1(A1, MaxColWdt)
Dim Hdr$: Hdr = WdtAyHdrLin(W)

If Not HidGrid Then Push Dry_Fmt, Hdr       '<=============
If BrkColIx >= 0 Then
    PushIAy Dry_Fmt, Dry_XIns_BrkLin(A1, BrkColIx, Hdr, W) '<=================
Else
    Dim Dr
    For Each Dr In A1
        PushI Dry_Fmt, Dr_Fmt(Dr, W, HidGrid)  '<======================================
    Next
End If
If Not HidGrid Then PushI Dry_Fmt, Hdr      '<=============
End Function

Function Dry_StrCellDry(Dry, ShwZer As Boolean) As Variant()
Dim Dr
For Each Dr In Dry
   Push Dry_StrCellDry, Dr_StrCellDr(Dr, ShwZer)
Next
End Function

Private Function Dr_StrCellDr(Dr, ShwZer As Boolean) As String()
Dim I
For Each I In Dr
    PushI Dr_StrCellDr, Val_StrCell(I, ShwZer)
Next
End Function
Private Function Val_StrCell(V, Optional ShwZer As Boolean) ' Convert V into a string in a cell
'CellStr is a string can be displayed in a cell
Select Case True
Case IsNumeric(V)
    If V = 0 Then
        If ShwZer Then
            Val_StrCell = "0"
        End If
    Else
        Val_StrCell = V
    End If
Case IsEmp(V):
Case IsArray(V)
    Dim N&: N = Sz(V)
    If N = 0 Then
        Val_StrCell = "*[0]"
    Else
        Val_StrCell = "*[" & N & "]" & V(0)
    End If
Case IsObject(V): Val_StrCell = TypeName(V)
Case Else:        Val_StrCell = V
End Select
End Function

Private Function Dry_XIns_BrkLin(Dry, BrkColIx%, Hdr$, W%()) As String()
Dim Dr, DrIx&, IsBrk As Boolean
Push Dry_XIns_BrkLin, Hdr
For Each Dr In Dry
    IsBrk = ZIsBrk(Dry, DrIx, BrkColIx)
    If IsBrk Then Push Dry_XIns_BrkLin, Hdr
    Push Dry_XIns_BrkLin, Dr_Fmt(Dr, W)
    DrIx = DrIx + 1
Next
End Function

Private Function ZIsBrk(Dry, DrIx&, BrkColIx%) As Boolean
If Sz(Dry) = 0 Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(Dry) Then Exit Function
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Function
ZIsBrk = True
End Function
