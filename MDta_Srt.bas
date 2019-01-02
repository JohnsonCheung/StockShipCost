Attribute VB_Name = "MDta_Srt"
Option Compare Binary
Option Explicit

Function Drs_XSrt(A As Drs, FF, Optional IsDes As Boolean) As Drs
Set Drs_XSrt = New_Drs(A.Fny, Dry_XSrt(A.Dry, Ay_Ix(A.Fny, FF), IsDes))
End Function

Function Dry_XSrt(Dry, Optional ColIxAy, Optional IsDes As Boolean) As Variant()
Dim Col: Col = DryCol(Dry, ColIxAy)
Dim Ix&(): Ix = Ay_XSrt_IntoIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, Dry(Ix(J))
Next
Dry_XSrt = O
End Function

Function Dt_XSrt(A As Dt, FF, Optional IsDes As Boolean) As Dt
Set Dt_XSrt = New_Dt(A.DtNm, A.Fny, Drs_XSrt(Dt_Drs(A), FF, IsDes).Dry)
End Function

