Attribute VB_Name = "MVb_Fs_Sel"
Option Compare Binary
Option Explicit

Function FfnSel$(A, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = A
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function


Function Pth_XSel$(Optional A$, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    If Not IsEmp(A) Then .InitialFileName = A
    .Show
    If .SelectedItems.Count = 1 Then
        Pth_XSel = Pth_XEns_Sfx(.SelectedItems(1))
    End If
End With
End Function

Private Sub Z_Pth_XSel()
GoTo ZZ
ZZ:
MsgBox FfnSel("C:\")
End Sub


Private Sub Z()
Z_Pth_XSel
MVb_Fs_Sel:
End Sub
