Attribute VB_Name = "MIde_Z_Pj_Rf"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Z_Pj_Rf."

Property Get CurPj_RfFfnAy() As String()
CurPj_RfFfnAy = Pj_RfFfnAy(CurPj)
End Property

Property Get CurPj_RfFmt() As String()
CurPj_RfFmt = Ay_XAlign_2T(Pj_RfLy(CurPj))
End Property

Property Get CurPj_RfLy() As String()
CurPj_RfLy = Pj_RfLy(CurPj)
End Property

Property Get CurPj_RfNy() As String()
CurPj_RfNy = PjRfNy(CurPj)
End Property

Function CvRf(A) As VBIDE.Reference
Set CvRf = A
End Function
Sub Pj_XSet_Rf_ByRfFfnAy(A As VBProject, B$())

End Sub
Sub Pj_XCpy_Rf(A As VBProject, ToPj As VBProject)
Pj_XSet_Rf_ByRfFfnAy ToPj, Pj_RfFfnAy(A)
End Sub

Function Pj_XHas_Rf(A As VBProject, RfNm)
Dim RF As VBIDE.Reference
For Each RF In A.References
    If RF.Name = RfNm Then Pj_XHas_Rf = True: Exit Function
Next
End Function
Function Pj_XHas_RfGUID(A As VBProject, RfGUID)
Pj_XHas_RfGUID = Itr_XHas_PrpVal(A.References, "GUID", RfGUID)
End Function

Function Pj_XHas_RfFfn(A As VBProject, RfFfn) As Boolean
Pj_XHas_RfFfn = Itr_XHas_PrpVal(A.References, "FullPath", RfFfn)
End Function

Function Pj_XHas_RfNm(A As VBProject, RfNm$) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.Name = RfNm Then Pj_XHas_RfNm = True: Exit Function
Next
End Function
Sub Pj_XSet_Rf_ByRfFil(A As VBProject, B$)

End Sub

Sub Pj_XSet_Rf_ByCfgFil(A As VBProject, B$)
Dim C As Dictionary: Set C = Ft_Dic(B)
Dim K
For Each K In C.Keys
    Pj_XSet_Rf_ByRfFfn A, K, C(K)
Next
End Sub

Function Pj_RfLy_FmCfgFil(A As VBProject) As String()
Const CSub$ = CMod & "PjReadRfCfg"
Dim B$: B = Pj_RfCfgFfn(A)
Ffn_XAss_Exist B, CSub, "Pj-Rf-Cfg-Fil"
Pj_RfLy_FmCfgFil = Ft_Ly(B)
End Function

Sub Pj_XBrw_Rf_FmPj(A As VBProject)
Stop '
'Ay_XBrw Pj_RfLy_FmPj(A)
End Sub

Sub Pj_XBrw_Rf_FmFil(A As VBProject)
Ay_XBrw Pj_RfLy_FmCfgFil(A)
End Sub

Function Pj_RfCfgFfn$(A As VBProject)
Pj_RfCfgFfn = Pj_SrcPth(A) & "PjRf.Cfg"
End Function
Sub XDmp_Rf_FmCurPj()
Pj_XDmp_Rf_FmPj CurPj
End Sub
Sub XDmp_Rf_FmFil()
Pj_XDmp_Rf_FmFil CurPj
End Sub
Sub Pj_XDmp_Rf_FmPj(A As VBProject)
D Pj_RfLy_FmPj(A)
End Sub
Function Pj_RfLy_FmPj(A As VBProject) As String()

End Function


Sub Pj_XDmp_Rf_FmFil(A As VBProject)
D Pj_RfLy_FmCfgFil(A)
End Sub

Function Pj_RfFfnAy(A As VBProject) As String()
Pj_RfFfnAy = ItrPrp_Sy(A.References, "FullPath")
End Function
Function Rf_Lin$(A As VBIDE.Reference)
With A
Rf_Lin = .Name & " " & .GUID & " " & .Major & " " & .Minor & " " & .FullPath
End With
End Function

Function Pj_RfLy(A As VBProject) As String()
Pj_RfLy = Ay_XAlign_4T(Rfs_Ly(A.References))
End Function

Function PjRfNmRfFfn$(A As VBProject, RfNm$)
PjRfNmRfFfn = Pj_Pth(A) & RfNm & ".xlam"
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = Itr_Ny(A.References)
End Function

Sub PjRmvRf(A As VBProject, RfNy0$)
Ay_XDoPX CvNy(RfNy0), "PjRmvRf__X", A
Pj_XSav A
End Sub

Private Sub PjRmvRf__X(A As VBProject, RfNm$)
If Pj_XHas_RfNm(A, RfNm) Then
    Debug.Print QQ_Fmt("Pj_XSet_Rf_ByRfFfn: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNmRfFfn(A, RfNm)
If Pj_XHas_RfFfn(A, RfFfn) Then
    Debug.Print QQ_Fmt("Pj_XSet_Rf_ByRfFfn: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub

Function RfFfn$(A As VBIDE.Reference)
On Error Resume Next
RfFfn = A.FullPath
End Function

Function Rfs_Ly(A As VBIDE.References) As String()
Dim R As VBIDE.Reference
For Each R In A
    PushI Rfs_Ly, Rf_Lin(R)
Next
End Function

Property Get RfNy() As String()
RfNy = CurPj_RfNy
End Property

Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function
Private Sub ZZ()
Dim A As Variant
Dim B As VBProject
Dim C$()
Dim D$
Dim E As Reference
Dim F As VBIDE.Reference
CvRf A
Pj_XSet_Rf_ByRfFfn B, A, A
Pj_XSet_Rf_ByRfFfnFfnAy B, C
End Sub

Private Sub Z()
End Sub
