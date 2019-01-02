Attribute VB_Name = "MDta_Wh"
Option Compare Binary
Option Explicit

Function Drs_XWh_FldEqV(A As Drs, F, EqVal) As Drs
Set Drs_XWh_FldEqV = New_Drs(A.Fny, Dry_XWh_(A.Dry, Ay_Ix(A.Fny, F), EqVal))
End Function

Function Drs_XWh_FFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
Set Drs_XWh_FFNe = New_Drs(Fny, Dry_XWh_CCNe(A.Dry, Ay_Ix(Fny, F1), Ay_Ix(Fny, F2)))
End Function

Function Drs_XWh_ColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = Ay_Ix(Fny, C)
Set Drs_XWh_ColEq = New_Drs(Fny, Dry_XWh_ColEq(A.Dry, Ix, V))
End Function

Function Drs_XWh_ColGt(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = Ay_Ix(Fny, C)
Set Drs_XWh_ColGt = New_Drs(Fny, Dry_XWh_ColGt(A.Dry, Ix, V))
End Function

Function Drs_XWh_NotIx(A As Drs, IxAy&()) As Drs
Dim ODry(), Dry()
    Dry = A.Dry
    Dim J&, I&
    For J = 0 To UB(A.Dry)
        If Not Ay_XHas(IxAy, J) Then
            PushI ODry, Dry(J)
        End If
    Next
Drs_XWh_NotIx = New_Drs(A.Fny, ODry)
End Function

Function Drs_XWh_NotRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not Ay_XHas(RowIxAy, J) Then
            Push O, Dry(J)
        End If
    Next
Set Drs_XWh_NotRowIxAy = New_Drs(A.Fny, O)
End Function

Function Drs_XWh_RowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O()
    Dim I, Dry()
    Dry = A.Dry
    For Each I In AyNz(RowIxAy)
        Push O, Dry(I)
    Next
Set Drs_XWh_RowIxAy = New_Drs(A.Fny, O)
End Function
