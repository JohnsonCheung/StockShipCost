Attribute VB_Name = "MDta_Fmt"
Option Compare Binary
Option Explicit

Function Drs_Fmt(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'If BrkColNm changed, insert a break line if BrkColNm is given
Dim Drs As Drs
    If HidIxCol Then
        Set Drs = A
    Else
        Set Drs = Drs_XAdd_RowIxCol(A)
    End If
Dim BrkColIx%
    BrkColIx = -1
    If Not HidIxCol Then
        BrkColIx = Ay_Ix(A.Fny, BrkColNm)
        If BrkColIx >= 0 Then
            BrkColIx = BrkColIx + 1 ' Need to increase by 1 due the Ix column is added
        End If
    End If
Dim Dry()
    Dry = Drs.Dry
    PushI Dry, Drs.Fny
Dim Ay$()
    Ay = Dry_Fmt(Dry, MaxColWdt, BrkColIx, ShwZer) '<== Will insert break line if BrkColIx>=0
Dim U&: U = UB(Ay)
Dim Hdr$: Hdr = Ay(U - 1)
Dim Lin$: Lin = Ay(U)
Drs_Fmt = Ay_XExl_LasNEle(AyInsAy(Ay, Array(Lin, Hdr)), 2)
PushI Drs_Fmt, Lin
End Function

Function Ds_Fmt(A As Ds, Optional MaxColWdt% = 100, Optional DtBrkColDicVbl$, Optional NoIxCol As Boolean) As String()
Push Ds_Fmt, "*Ds " & A.DsNm & " " & String(10, "=")
Dim Dic As Dictionary
    Set Dic = NewDicVBL(DtBrkColDicVbl)
Dim J%, Dt As Dt, BrkColNm$, DtAy() As Dt
DtAy = A.DtAy
For J = 0 To UB(DtAy)
    Set Dt = DtAy(J)
    If Dic.Exists(Dt.DtNm) Then BrkColNm = Dic(Dt.DtNm) Else BrkColNm = ""
    PushAy Ds_Fmt, DtFmt(Dt, MaxColWdt, BrkColNm, NoIxCol)
Next
End Function

Function DtFmt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
Push DtFmt, "*Tbl " & A.DtNm
PushAy DtFmt, Drs_Fmt(Dt_Drs(A), MaxColWdt, BrkColNm, ShwZer, HidIxCol)
End Function

Private Sub Z_Ds_Fmt()
Dim A As Ds, MaxColWdt%, DtBrkLinMapStr$, NoIxCol As Boolean
Set A = Samp_Ds
GoSub Tst
Exit Sub
Tst:
    Act = Ds_Fmt(A, MaxColWdt, DtBrkLinMapStr, NoIxCol)
    Brw Act: Stop
    C
    Return
End Sub

Private Sub Z_DtFmt()
Dim A As Dt, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
Set A = Samp_Dt1
'Ept = Z_DtFmtEpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = DtFmt(A, MaxColWdt, BrkColNm, ShwZer)
    C
    Return
End Sub

Private Sub ZZ()
Dim A As Drs
Dim B%
Dim C$
Dim D As Boolean
Dim E As Ds
Dim F As Dt
Drs_Fmt A, B, C, D, D
Ds_Fmt E, B, C, D
DtFmt F, B, C, D, D
End Sub

Private Sub Z()
Z_Ds_Fmt
Z_DtFmt
End Sub
