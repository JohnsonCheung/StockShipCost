Attribute VB_Name = "MDta_X_Dt"
Option Compare Binary
Option Explicit

Sub Dt_XBrw(A As Dt, Optional Fnn$)
Ay_XBrw DtFmt(A), Fnn
End Sub
Sub DtAy_XAss_DupNm(A() As Dt, Fun$)
Ay_XAss_Dup DtAy_DtNy(A), Fun
End Sub
Function DtAy_DtNy(A() As Dt) As String()
Dim I
For Each I In AyNz(A)
    PushI DtAy_DtNy, CvDt(I).DtNm
Next
End Function
Function DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(AyQuoteDbl(A.Fny))
For Each Dr In A.Dry
   Push O, QQ_FmtAv(QQStr, Dr)
Next
End Function

Function Dt_XDrp_Col(A As Dt, FF, Optional DtNm$) As Dt
Dim O As Drs: Set O = Drs_XDrp_Col(Dt_Drs(A), FF)
Set Dt_XDrp_Col = New_Dt(Dft(DtNm, A.DtNm), O.Fny, O.Dry)
End Function

Function Dt_Drs(A As Dt) As Drs
Set Dt_Drs = New_Drs(A.Fny, A.Dry)
End Function

Sub Dt_XDmp(A As Dt)
Ay_XDmp DtFmt(A)
End Sub
Property Get EmpDtAy() As Dt()
End Property

Function DtIsEmp(A As Dt) As Boolean
DtIsEmp = Sz(A.Dry) = 0
End Function

Function DtReOrd(A As Dt, ColLvs$) As Dt
Dim ReOrdFny$(): ReOrdFny = Ssl_Sy(ColLvs)
Dim IxAy&(): IxAy = Ay_IxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
DtReOrd.DtNm = A.DtNm
Set DtReOrd = New_Drs(OFny, ODry)
End Function
Function New_Dt(DtNm, Fny0, Dry()) As Dt
Dim O As New Dt
Set New_Dt = O.Init(DtNm, Fny0, Dry)
End Function

Function CvDt(A) As Dt
Set CvDt = A
End Function
