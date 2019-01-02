Attribute VB_Name = "MDao_Att_Op_Dlt"
Option Compare Binary
Option Explicit

Sub DbDrpAtt(A As Database, Att)
A.Execute QQ_Fmt("Delete * from Att where AttNm='?'", Att)
End Sub


Sub AttyDrp(Atty0)
DbDrpAtty CurDb, Atty0
End Sub




Sub DbDrpAtty(A As Database, Atty0)
Ay_XDoPX CvNy(Atty0), "DbDrpAtt", A
End Sub


Sub AttDrp(Att)
DbDrpAtt CurDb, Att
End Sub


