Attribute VB_Name = "MIde_Cmd_Lis_Src"
Option Compare Binary
Option Explicit
Sub LisSrc(Patn$)
CurPjLisSrc Patn
End Sub
Sub CurPjLisSrc(Patn$)
PjLisSrc CurPj, Patn
End Sub
Sub PjLisSrc(A As VBProject, Patn$)
Dim R As RegExp, O$()
Set R = New_Re(Patn)
Dim C As VBComponent
For Each C In A.VBComponents
    PushIAy O, CmpReSrc(C, R)
Next
D Ay_Fmt(O, ", '")
End Sub

Function MdNmLnoGoStr$(MdDNm$, Lno&)
MdNmLnoGoStr = QQ_Fmt("MdNmLnoGo ""?"",?", MdDNm, Lno)
End Function

Function CmpReSrc(A As VBComponent, R As RegExp) As String()
Dim Md As CodeModule, L$, J&, N$
Set Md = A.CodeModule
N = Md_Nm(Md)
For J = 1 To Md.CountOfLines
    L = Md.Lines(J, 1)
    If R.Test(L) Then
        PushI CmpReSrc, MdNmLnoGoStr(N, J) & " ' " & L
    End If
Next
End Function

