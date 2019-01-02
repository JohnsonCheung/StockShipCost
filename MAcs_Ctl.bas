Attribute VB_Name = "MAcs_Ctl"
Option Compare Binary
Option Explicit

Sub TxtbSelPth(A As Access.TextBox)
Dim R$
R = Pth_XSel(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub
Sub CmdTurnOffTabStop(AcsCtl As Access.Control)
Dim A As Access.Control
Set A = AcsCtl
If Not XHas_Pfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTgl(A): CvTgl(A).TabStop = False
End Select
End Sub

Sub FrmSetCmdNotTabStop(A As Access.Form)
Itr_XDo A.Controls, "CmdTurnOffTabStop"
End Sub

Function CvCtl(A) As Access.Control
Set CvCtl = A
End Function

Function CvBtn(A) As Access.CommandButton
Set CvBtn = A
End Function


Function CvTgl(A) As Access.ToggleButton
Set CvTgl = A
End Function
