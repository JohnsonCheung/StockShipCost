Attribute VB_Name = "MVb_Ay_Seq"
Option Compare Binary
Option Explicit


Function SeqInto(FmNum, ToNum, OAy)
Dim O&()
ReDim OAy(Abs(FmNum - ToNum))
Dim J&, I&
If ToNum > FmNum Then
    For J = FmNum To ToNum
        OAy(I) = J
        I = I + 1
    Next
Else
    For J = ToNum To FmNum Step -1
        OAy(I) = J
        I = I + 1
    Next
End If
End Function

Function SeqOfInt(FmNum%, ToNum%) As Integer()
SeqOfInt = SeqInto(FmNum, ToNum, EmpIntAy)
End Function

Function SeqOfLng(FmNum&, ToNum&) As Long()
SeqOfLng = SeqInto(FmNum, ToNum, EmpLngAy)
End Function

