Attribute VB_Name = "MAdo_Rs"
Option Compare Binary
Option Explicit
Const Samp_Fb_DutyPrepare$ = ""

Function ARsF_Into(A As ADODB.Recordset, OInto, Optional Col = 0)
ARsF_Into = Ay_XCln(OInto)
With A
    While Not .EOF
        PushI ARsF_Into, Nz(.Fields(Col).Value, Empty)
        .MoveNext
    Wend
    .Close
End With
End Function

Function ARs_Drs(A As ADODB.Recordset) As Drs
Set ARs_Drs = New_Drs(ARs_Fny(A), ARs_Dry(A))
End Function

Function ARs_Dry(A As ADODB.Recordset) As Variant()
While Not A.EOF
    PushI ARs_Dry, AFds_Dr(A.Fields)
    A.MoveNext
Wend
End Function

Function ARs_Fny(A As ADODB.Recordset) As String()
ARs_Fny = AFds_Fny(A.Fields)
End Function

Function ARs_IntAy(A As ADODB.Recordset, Optional Col = 0) As Integer()
ARs_IntAy = ARsF_Into(A, EmpIntAy, Col)
End Function

Function ARs_Sy(A As ADODB.Recordset, Optional Col = 0) As String()
ARs_Sy = ARsF_Into(A, EmpSy, Col)
End Function

Private Sub Z_ARs_Dry()
Dim S$
Const Q$ = "Select * from KE24"
S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
Dry_XBrw ARs_Dry(Cnq_ARs(Fb_Cn(Samp_Fb_DutyPrepare), Q))
End Sub



Private Sub Z()
Z_ARs_Dry
End Sub

Private Sub ZZ()
Dim A As ADODB.Recordset
Dim B
Dim XX
ARs_Drs A
ARs_Dry A
ARs_Fny A
End Sub
