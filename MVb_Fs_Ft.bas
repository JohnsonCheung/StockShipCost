Attribute VB_Name = "MVb_Fs_Ft"
Option Compare Binary
Option Explicit

Sub Ft_XBrw(A)
Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub

Function Ft_Dic(A) As Dictionary
Set Ft_Dic = New_Dic_LY(Ft_Ly(A))
End Function

Function Ft_XEns$(A)
If Not Ffn_Exist(A) Then Str_XWrt "", A
Ft_XEns = A
End Function

Function Ft_Lines$(A)
If Ffn_Sz(A) <= 0 Then Exit Function
Ft_Lines = Fso.GetFile(A).OpenAsTextStream.ReadAll
End Function

Function Ft_Ly(A) As String()
Ft_Ly = SplitLines(Ft_Lines(A))
End Function

Function Ft_Fno_ForApp%(A)
Dim O%: O = FreeFile(1)
Open A For Append As #O
Ft_Fno_ForApp = O
End Function

Function Ft_Fno_ForInp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
Ft_Fno_ForInp = O
End Function

Function Ft_Fno_ForOup%(A)
Dim O%: O = FreeFile(1)
Open A For Output As #O
Ft_Fno_ForOup = O
End Function

Sub Ft_XRmv_Fst4Lines(A)
' This approach will preserve the Class-Attribute during import/export
Dim A1$: A1 = Fso.GetFile(A).OpenAsTextStream.ReadAll
Dim A2$: A2 = Left(A1, 55)
Dim A3$: A3 = Mid(A1, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If A2 <> B1 Then Stop
Fso.CreateTextFile(A, True).Write A3
End Sub
