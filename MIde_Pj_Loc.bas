Attribute VB_Name = "MIde_Pj_Loc"
Option Explicit

Function Pj_Src(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushIAy Pj_Src, Md_Src(C.CodeModule)
Next
End Function
Function CurPj_LocLy_PATN(Patn$) As String()
CurPj_LocLy_PATN = Pj_LocLy_PATN(CurPj, Patn)
End Function

Function Pj_LocLy_PATN(A As VBProject, Patn$) As String()
Pj_LocLy_PATN = Ay_XWh_Patn(Pj_Src(A), Patn)
End Function

Function CurPj_LocLy(Re_Or_Patn) As String()

End Function

Function Pj_LocLy(A As VBProject, Re As RegExp) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy Pj_LocLy, Md_LocLy(C.CodeModule, Re)
Next
End Function

Function Md_LocLy(A As CodeModule, Re As RegExp) As String()
End Function
