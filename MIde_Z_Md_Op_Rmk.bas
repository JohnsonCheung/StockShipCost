Attribute VB_Name = "MIde_Z_Md_Op_Rmk"
Option Compare Binary
Option Explicit
Sub XRmk_Md()
Md_XRmk CurMd
End Sub
Sub XUnRmk_Md()
Md_XUnRmk CurMd
End Sub
Sub XRmk_AllMd()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In CurPj_MdAy
    If Md.Name <> "LibIdeRmkMd" Then
        If Md_XRmk(CvMd(I)) Then
            NRmk = NRmk + 1
        Else
            Skip = Skip + 1
        End If
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub


Function Md_IsAllRmked(A As CodeModule) As Boolean
Dim J%, L$
For J = 1 To A.CountOfLines
    If Left(A.Lines(J, 1), 1) <> "'" Then Exit Function
Next
Md_IsAllRmked = True
End Function


Sub XUnRmk_AllMd()
Dim I, Md As CodeModule
Dim NUnRmk%, Skip%
For Each I In CurPj_MdAy
    Set Md = I
    If Md_XUnRmk(Md) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

Private Function Md_XRmk(A As CodeModule) As Boolean
Debug.Print "Rmk " & A.Parent.Name,
If Md_IsAllRmked(A) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To A.CountOfLines
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
Md_XRmk = True
End Function

Private Function Md_XUnRmk(A As CodeModule) As Boolean
Debug.Print "UnRmk " & A.Parent.Name,
If Not Md_IsAllRmked(A) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
Md_XUnRmk = True
End Function
