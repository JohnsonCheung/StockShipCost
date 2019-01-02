Attribute VB_Name = "MVb_Er_Prompt"
Option Compare Binary
Option Explicit
Sub Done()
MsgBox "Done"
End Sub
Sub Er_XHalt(Er$())
If Sz(Er) = 0 Then Exit Sub
Ay_XBrw Er
XHalt
End Sub

