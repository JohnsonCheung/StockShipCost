Attribute VB_Name = "MDao_Z_Fds"
Option Compare Binary
Option Explicit
Function Fds_IsEq(A As DAO.Fields, B As DAO.Fields) As Boolean
Stop '
End Function

Sub Fds_XSet_SqRow(A As DAO.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
DrSetSqRow Fds_Dr(A), Sq, R, NoTxtSngQ
End Sub

Function Fds_Csv$(A As DAO.Fields)
Fds_Csv = AyCsv(Itr_Vy(A))
End Function
Function NzEmpty(A)
If IsNull(A) Then Exit Function
Asg A, NzEmpty
End Function
Function Fds_Dr(A As DAO.Fields) As Variant()
Dim F As DAO.Field
For Each F In A
    PushI Fds_Dr, NzEmpty(F.Value)
Next
End Function

Function Fds_Fny(A As Fields) As String()
Fds_Fny = Itr_Ny(A)
End Function

Function Fds_XHas_FldNm(A As DAO.Fields, F) As Boolean
Fds_XHas_FldNm = Itr_XHas_Nm(A, F)
End Function


Function Fds_Dr_ByFF(A As DAO.Fields, FF) As Variant()
Fds_Dr_ByFF = Fds_Vy_ByFF(A, FF)
End Function

Function Fds_Vy(A As DAO.Fields) As Variant()
Fds_Vy = Itr_Vy(A)
End Function

Function Fds_Vy_ByFF(A As DAO.Fields, FF) As Variant()
Dim O(), J%, F
Dim Fny$()
    Fny = FF_Fny(FF)
ReDim O(UB(Fny))
For Each F In Fny
    O(J) = A(F).Value
    J = J + 1
Next
Fds_Vy_ByFF = O
End Function

Private Sub Z_Fds_Dr()
Dim Rs As DAO.Recordset, Dry()
Set Rs = Fb_Db(Samp_Fb_ShpRate).OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        Push Dry, Fds_Dr(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
Brw Dry_Fmt(Dry)
End Sub

Private Sub Z_Fds_Vy()
Dim Rs As DAO.Recordset, Vy()
'Set Rs = CurDb.OpenRecordset("Select * from SkuB")
With Rs
    While Not .EOF
        Vy = Fds_Vy(Rs)
        Debug.Print JnComma(Vy)
        .MoveNext
    Wend
    .Close
End With
End Sub



Private Sub Z()
Z_Fds_Dr
Z_Fds_Vy
MDao_Z_Fds:
End Sub
