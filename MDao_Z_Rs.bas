Attribute VB_Name = "MDao_Z_Rs"
Option Compare Binary
Option Explicit
Private Sub ZZ_Rs_Asg()
Dim Y As Byte, M As Byte
Rs_Asg Tbl_Rs("YM"), Y, M
Stop
End Sub
Function CvRs(A) As DAO.Recordset
Set CvRs = A
End Function
Function Rs_XHas_Rec(A As DAO.Recordset) As Boolean
Rs_XHas_Rec = Not Rs_IsNoRec(A)
End Function

Function Rs_IsNoRec(A As DAO.Recordset) As Boolean
If A.EOF Then Exit Function
If A.BOF Then Exit Function
Rs_IsNoRec = True
End Function

Sub Rs_Asg(A As DAO.Recordset, ParamArray OAp())
Dim F As DAO.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Function Rs_Ay(A As DAO.Recordset, Optional F0) As Variant()
Rs_Ay = Rs_AyInto(A, Rs_Ay, F0)
End Function

Function Rs_AyInto(A As DAO.Recordset, OInto, Optional F0)
Dim O: O = OInto: Erase O
Dim F
F = DftF0(F0)
With A
    If .EOF Then Rs_AyInto = O: Exit Function
    .MoveFirst
    While Not .EOF
        Push O, .Fields(F).Value
        .MoveNext
    Wend
End With
Rs_AyInto = O
End Function

Sub Rs_XBrw(A As DAO.Recordset)
Drs_XBrw Rs_Drs(A)
End Sub

Sub Rs_XBrw_zSingleRec(A As DAO.Recordset)
Ay_XBrw Rs_Ly_SINGLE_REC(A)
End Sub

Sub RsDlt(A As DAO.Recordset)
With A
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub
Sub Rs_XClr(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

Function RsCsv$(A As DAO.Recordset)
RsCsv = Fds_Csv(A.Fields)
End Function

Function Rs_CsvLy(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "Rs_CsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    For Each F In A.Fields
        Dr(I) = CvCsv(F.Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
Rs_CsvLy = O
End Function

Function Rs_CsvLyByFny0(A As DAO.Recordset, Fny0) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = CvNy(Fny0)
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "Rs_CsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    Set Flds = A.Fields
    For Each F In Fny
        Dr(I) = CvCsv(Flds(F).Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
Rs_CsvLyByFny0 = O
End Function
Function Rs_Fmt(A As Recordset)
Rs_Fmt = Dic_Fmt(Rs_Dic_OneRec(A), InclValTy:=True)
End Function
Sub Rs_XDmp_OneRec(A As Recordset)
D Rs_Fmt(A)
A.MoveFirst
End Sub

Sub Rs_XDmpByFny0(A As Recordset, Fny0)
Ay_XDmp Rs_CsvLyByFny0(A, Fny0)
A.MoveFirst
End Sub

Function Rs_Dr(A As DAO.Recordset) As Variant()
Rs_Dr = Fds_Dr(A.Fields)
End Function

Function Rs_Dr_ByFF(A As DAO.Recordset, FF) As Variant()
Rs_Dr_ByFF = Fds_Dr_ByFF(A.Fields, FF)
End Function

Function Rs_Drs(A As DAO.Recordset) As Drs
Set Rs_Drs = New_Drs(Rs_Fny(A), Rs_Dry(A))
End Function

Function Rs_Dry(A As DAO.Recordset, Optional InclFldNm As Boolean) As Variant()
If Not Rs_IsNoRec(A) Then Exit Function
If InclFldNm Then
    PushI Rs_Dry, Rs_Fny(A)
End If
With A
    .MoveFirst
    While Not .EOF
        PushI Rs_Dry, Fds_Dr(.Fields)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function

Function Rs_Fny(A As DAO.Recordset) As String()
Rs_Fny = Itr_Ny(A.Fields)
End Function

Function Rs_XHas_FldEqV(A As DAO.Recordset, F$, EqVal) As Boolean
With A
    If .BOF Then
        If .EOF Then Exit Function
    End If
    .MoveFirst
    While Not .EOF
        If .Fields(F) = EqVal Then Rs_XHas_FldEqV = True: Exit Function
        .MoveNext
    Wend
End With
End Function

Function Rs_IntAy(A As DAO.Recordset, Optional F) As Integer()
Rs_IntAy = Rs_AyInto(A, Rs_IntAy)
End Function

Function Rs_IsBrk(A As DAO.Recordset, GpKK, LasVy()) As Boolean
Rs_IsBrk = Not Ay_IsEq(Rs_Dr_ByFF(A, GpKK), LasVy)
End Function

Function Rs_Lin$(A As DAO.Recordset, Optional Sep$ = " ")
Rs_Lin = Join(Rs_Dr(A), Sep)
End Function

Function Rs_LngAy(A As DAO.Recordset, Optional FldNm$) As Long()
Rs_LngAy = Rs_AyInto(A, FldNm, Rs_LngAy)
End Function

Function Rs_Ly(A As DAO.Recordset, Optional Sep$ = " ") As String()
Dim O$()
With A
    Push O, Join(Rs_Fny(A), Sep)
    While Not .EOF
        Push O, Rs_Lin(A, Sep)
        .MoveNext
    Wend
End With
Rs_Ly = O
End Function

Function Rs_Ly_SINGLE_REC(A As DAO.Recordset)
Rs_Ly_SINGLE_REC = NyAv_Ly(Rs_Fny(A), Rs_Dr(A))
End Function

Function RsMovFst(A As DAO.Recordset) As DAO.Recordset
A.MoveFirst
Set RsMovFst = A
End Function

Function RsNRec&(A As DAO.Recordset)
Dim O&
With A
    .MoveFirst
    While Not .EOF
        O = O + 1
        .MoveNext
    Wend
    .MoveFirst
End With
RsNRec = O
End Function


Sub Rs_XPut_Sq(A As DAO.Recordset, Sq, R&, Optional NoTxtSngQ As Boolean)
Fds_XSet_SqRow A.Fields, Sq, R, NoTxtSngQ
End Sub

Function Rs_Sq(A As DAO.Recordset, Optional InclFldNm As Boolean) As Variant()
Rs_Sq = Dry_Sq(Rs_Dry(A, InclFldNm))
End Function

Function Rs_StrCol_FstCol(A As DAO.Recordset) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(0).Value
        .MoveNext
    Wend
End With
Rs_StrCol_FstCol = O
End Function

Function Rs_Stru$(A As DAO.Recordset)
Dim O$(), F As DAO.Field2
For Each F In A.Fields
    PushI O, Fd_Str(F)
Next
Rs_Stru = JnCrLf(O)
End Function
Function Nz(A, Optional B = "")
Nz = IIf(IsNull(A), B, A)
End Function
Function RsF_Into(A As Recordset, F, OInto)
RsF_Into = Ay_XCln(OInto)
While Not A.EOF
    PushI RsF_Into, Nz(A(F).Value, Empty)
    A.MoveNext
Wend
End Function

Function Rs_Sy(A As DAO.Recordset, Optional F = 0) As String()
Rs_Sy = RsF_Into(A, F, EmpSy)
End Function

Function Rs_Val(A As DAO.Recordset)
If Rs_IsNoRec(A) Then Rs_Val = A.Fields(0).Value
End Function

Property Let RsF_Val(A As DAO.Recordset, F, V)
With A
    .Edit
    .Fields(F).Value = V
    .Update
End With
End Property

Property Get RsF_Val(A As DAO.Recordset, F)
With A
    If .EOF Then Exit Property
    If .BOF Then Exit Property
    RsF_Val = .Fields(F).Value
End With
End Property

Function Rs_TimSzDotStr$(A As DAO.Recordset)
If A.Fields(0).Type <> DAO.dbDate Then Stop
If A.Fields(1).Type <> DAO.dbLong Then Stop
If Rs_XHas_Rec(A) Then Exit Function
Rs_TimSzDotStr = Dte_DTim(A.Fields(0).Value) & "." & A.Fields(1).Value
End Function
Sub Ap_XIns_Rs(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
Dr_XIns_Rs Dr, Rs
End Sub

Sub Dr_XIns_Rs(A, Rs As DAO.Recordset)
Rs.AddNew
Dr_XSet_Rs A, Rs
Rs.Update
End Sub

Sub Ap_XUpd_Rs(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
Dr_XUpd_Rs Dr, Rs
End Sub

Sub Dr_XSet_Rs(Dr, Rs As DAO.Recordset)
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        Rs(J).Value = Rs(J).DefaultValue
    Else
        Rs(J).Value = V
    End If
    J = J + 1
Next
End Sub


Sub Dr_XUpd_Rs(A, Rs As DAO.Recordset)
If Sz(A) = 0 Then Exit Sub
Rs.Edit
Dr_XSet_Rs A, Rs
Rs.Update
End Sub
