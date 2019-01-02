Attribute VB_Name = "MVb_Lin_Term_Asg"
Option Compare Binary
Option Explicit

Sub Lin_2TRstAsg(A, OT1, OT2, ORst)
AyAsg Lin_2TRst(A), OT1, OT2, ORst
End Sub

Sub Lin_3TRstAsg(A, OT1, OT2, OT3, ORst)
AyAsg Lin_3TRst(A), OT1, OT2, OT3, ORst
End Sub

Sub Lin_4TRstAsg(A, OT1, OT2, OT3, OT4, ORst)
AyAsg Lin_4TRst(A), OT1, OT2, OT3, OT4, ORst
End Sub

Sub Lin_TRstAsg(A, OT1, ORst)
AyAsg LinN_NTermRst(A, 1), OT1, ORst
End Sub

Sub Lin_TTAsg(A, O1, O2)
AyAsg Lin_TT(A), O1, O2
End Sub
