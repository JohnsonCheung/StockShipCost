Attribute VB_Name = "MAcs_Ctl_TBox"
Option Explicit
Sub TBoxSet(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = Lines_LasNLin(A.Value & CrLf & Now & " " & Msg, 5)
A.Value = B
DoEvents
End Sub


