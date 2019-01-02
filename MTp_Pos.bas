Attribute VB_Name = "MTp_Pos"
Option Compare Binary
Option Explicit
Function TpPos_FmtStr$(A As TpPos)
Dim O$
With A
    Select Case .Ty
    Case ePosRCC
        O = QQ_Fmt("RCC(? ? ?) ", .R1, .C1, .C2)
    Case ePosRR
        O = QQ_Fmt("RR(? ?) ", .R1, .R2)
    Case ePosR
        O = QQ_Fmt("R(?)", .R1)
    Case Else
        'XThw CSub TpPos_FmtStr", "Invalid {TpPos}", A.Ty
    End Select
End With
End Function
