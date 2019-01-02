Attribute VB_Name = "MIde_Mth_Nm_DDNm"
Option Compare Binary
Option Explicit
Function Lin_MthDDNm$(A$)
'MthDDNm is a string of Nm.ShtKd.ShtMdy or Blank
Dim C$(): C = Lin_MthNmBrk(A): If Sz(C) = 0 Then Exit Function
Lin_MthDDNm = C(0) & "." & MthTy_MthShtTy(C(1)) & "." & MthMdy_MthShtMdy(C(2))
End Function

Function Md_MthDDNmDic(A As CodeModule) As Dictionary
Set Md_MthDDNmDic = DicAddKeyPfx(Src_MthNmDic(Md_Src(A)), Md_DNm(A) & ".")
End Function

Function Src_MthDDNmDic(A$()) As Dictionary
Dim Ix
Set Src_MthDDNmDic = New Dictionary
Src_MthDDNmDic.Add "*Dcl", Src_DclLines(A)
For Each Ix In AyNz(Src_MthIxAy(A))
    Src_MthDDNmDic.Add Lin_MthDDNm(A(Ix)), SrcMthIx_MthLines_WithTopRmk(A, CLng(Ix))
Next
End Function

Function MthDDNm_Mth(A$) As Mth
Dim M As CodeModule
Dim MthNm$
    Dim Ay$()
    Ay = SplitDot(A)
    Select Case Sz(Ay)
    Case 1: Set M = CurMd: MthNm = Ay(0)
    Case 2: Set M = Md(Ay(0)): MthNm = Ay(1)
    Case 3: Set M = Md(Ay(0) & "." & Ay(1)): MthNm = Ay(2)
    Case Else: XThw CSub, "MthDDNm should have 1 to 2 dot", "MthDDNm", A
    End Select
Set MthDDNm_Mth = New_Mth(M, MthNm)
End Function

Function MthDDNm_IsSel(A, B As WhMth) As Boolean
MthDDNm_IsSel = MthNmBrk_IsSel(SplitDot(A), B)
End Function
