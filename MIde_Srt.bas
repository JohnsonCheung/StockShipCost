Attribute VB_Name = "MIde_Srt"
Option Compare Binary
Option Explicit

Property Get CurMd_SrtLines$()
CurMd_SrtLines = Md_SrtLines(CurMd)
End Property

Function Lin_MthSrtKey$(A)
Dim L$, Mdy$, Ty$, Nm$
L = A
Mdy = XShf_MthMdy(L)
Ty = XShf_MthShtTy(L): If Ty = "" Then Exit Function
Nm = XTak_Nm(L)
Lin_MthSrtKey = Mth3Nm_SrtKey(Mdy, Ty, Nm)
End Function

Sub Md_XBrw_SrtRpt(A As CodeModule)
Dim Old$: Old = Md_BdyLines(A)
Dim NewLines$: NewLines = Md_SrtLines(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print Md_Nm(A), O
End Sub

Sub Md_XSrt(A As CodeModule)
Dim Nm$: Nm = Md_Nm(A)
Debug.Print "Sorting: "; XAlignL(Nm, 20); " ";
Dim LinesN$: LinesN = Md_SrtLines(A)
Dim LinesO$: LinesO = Md_Lines(A)
'Exit if same
    If LinesO = LinesN Then
        Debug.Print "<== Same"
        Exit Sub
    End If
'Delete
    Debug.Print QQ_Fmt("<--- Deleted (?) lines", A.CountOfLines);
    Md_XClr A, IsSilent:=True
'Add sorted lines
    A.AddFromString LinesN
    Md_XRmv_EndBlankLin A
    Debug.Print "<----Sorted Lines added...."
End Sub

Function Md_SrtLines$(A As CodeModule)
Md_SrtLines = Src_SrtLines(Md_Src(A))
End Function
Property Get CurMd_Src_SORTED() As String()
CurMd_Src_SORTED = Md_Src_SORTED(CurMd)
End Property
Function Md_Src_SORTED(A As CodeModule) As String()
Md_Src_SORTED = Src_SrtLy(Md_Src(A))
End Function

Function MthDDNm_MthSrtKey$(A) ' MthDDNm is Nm.Ty.Mdy
If A = "*Dcl" Then MthDDNm_MthSrtKey = "*Dcl": Exit Function
Dim B$(): B = SplitDot(A): If Sz(B) <> 3 Then XThw CSub, "MthDDNm should be 2-Dot", "MthDDNm", A
Dim Mdy$, Ty$, Nm$
AyAsg B, Nm, Ty, Mdy
MthDDNm_MthSrtKey = Mth3Nm_SrtKey(Mdy, Ty, Nm)
End Function

Function Mth3Nm_SrtKey$(Mdy$, Ty$, Nm$)
Dim P% 'Priority
    Select Case True
    Case XHas_Pfx(Nm, "Init"): P = 1
    Case Nm = "Z":           P = 9
    Case Nm = "ZZ":          P = 8
    Case XHas_Pfx(Nm, "Z_"):   P = 7
    Case XHas_Pfx(Nm, "ZZ_"):  P = 6
    Case XHas_Pfx(Nm, "Z"):    P = 5
    Case Else:               P = 2
    End Select
Mth3Nm_SrtKey = P & ":" & Nm & ":" & Ty & ":" & Mdy
End Function

Sub Pj_XSrt(A As VBProject)
Dim M As CodeModule, I, Ay() As CodeModule
If Sz(Ay) = 0 Then Exit Sub
For Each I In Ay
    Md_XSrt CvMd(I)
Next
End Sub

Function Src_SrtDic(A$()) As Dictionary
Dim D As Dictionary, K
Set D = Src_MthDDNmDic(A)
Dim O As New Dictionary
    For Each K In D
        O.Add MthDDNm_MthSrtKey(K), D(K)
    Next
Set Src_SrtDic = Dic_XSrt(O)
End Function

Function Src_SrtLines$(A$())
Src_SrtLines = JnDblCrLf(Src_SrtDic(A).Items)
End Function

Function Src_SrtLy(A$()) As String()
Src_SrtLy = SplitCrLf(Src_SrtLines(A))
End Function

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const MdNm$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Md_Src(Md(MdNm))
B = Src_SrtLy(A)
A1 = Src_DclLy(A)
B1 = Src_DclLy(B)
Stop
End Sub

Private Sub Z_Lin_MthSrtKey()
GoTo ZZ
Dim A$
'
Ept = "2:Lin_MthSrtKey:Function:": A = "Function Lin_MthSrtKey$(A)": GoSub Tst
Ept = "2:YYA:Function:":          A = "Function YYA()":            GoSub Tst
Exit Sub
Tst:
    Act = Lin_MthSrtKey(A)
    C
    Return
ZZ:
    Dim Ay1$(): Ay1 = Src_MthLinAy(CurSrc)
    Dim Ay2$(), L: For Each L In AyNz(Ay1): PushI Ay2, Lin_MthSrtKey(L): Next
    S1S2Ay_XBrw AyabS1S2Ay(Ay2, Ay1)
End Sub

Private Sub Z_MthDDNm_MthSrtKey()
GoSub X0
GoSub X1
Exit Sub
X0:
    Dim Ay1$(): Ay1 = Src_MthDDNy(CurSrc)
    Dim Ay2$(): Ay2 = AyMap_Sy(Ay1, "MthDDNm_MthSrtKey")
    Stop
    S1S2Ay_XBrw AyabS1S2Ay(Ay2, Ay1)
    Return
X1:
    Const A$ = "YYA.Fun."
    Debug.Print MthDDNm_MthSrtKey(A)
    Return
End Sub

Private Sub Z_Src_SrtLy()
Brw Src_SrtLines(CurSrc)
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As CodeModule
Dim C$
Dim D As VBProject
Dim E$()
Lin_MthSrtKey A
Md_XBrw_SrtRpt B
Md_XSrt B
Md_Src_SORTED B
MthDDNm_MthSrtKey A
Mth3Nm_SrtKey C, C, C
Pj_XSrt D
Src_SrtDic E
Src_SrtLines E
Src_SrtLy E
End Sub

Private Sub Z()
Z_Lin_MthSrtKey
Z_MthDDNm_MthSrtKey
Z_Src_SrtLy
End Sub
