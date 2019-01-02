Attribute VB_Name = "MIde_Mth_Rmk_StopLin"
Option Compare Database

Sub MdLno_XIns_StopLin_IfNeed(A As CodeModule, Lno&)
If A.Lines(Lno, 1) = "Stop '" Then Exit Sub
A.InsertLines Lno, "Stop '"
End Sub

Sub MdLno_XRmv_StopLin_IfAny(A As CodeModule, Lno&)
If A.Lines(Lno, 1) <> "Stop '" Then Exit Sub
A.DeleteLines Lno, 1
End Sub
