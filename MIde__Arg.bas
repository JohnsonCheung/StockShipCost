Attribute VB_Name = "MIde__Arg"
Option Compare Binary
Option Explicit

Function LinArg$(A)
If Lin_IsMth(A) Then
    LinArg = XTak_BetBkt(A)
End If
End Function

Function LinArgAy(A) As String()
LinArgAy = AyTrim(SplitComma(LinArg(A)))
End Function

Function PjArgAy(A As VBProject) As String()
Dim O$(), L
For Each L In AyNz(Pj_MthLinAy(A))
'    If L = "Std.MDao_Chk_Col.Function Dbt_XChk_Col_ByLnkColVbl(A As Database, T, LnkColVbl$) As String()" Then Stop
'    If Sz(LinArgAy(L)) > 0 Then Stop
    PushIAy O, LinArgAy(L)
Next
PjArgAy = Ay_XWh_DistFmt(O)
End Function

Private Sub Z_PjArgAy()
Brw AyQSrt(PjArgAy(CurPj))
End Sub
