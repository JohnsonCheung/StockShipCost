Attribute VB_Name = "MVb_Fs_CurDir"
Option Compare Binary
Option Explicit
Function CurFnAy(Optional Spec$ = "*") As String()
CurFnAy = Pth_FnAy(CurDir, Spec)
End Function

Function CurSubFdrAy(Optional Spec$ = "*") As String()
CurSubFdrAy = Pth_FdrAy(CurDir)
End Function
