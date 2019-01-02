Attribute VB_Name = "MVb_Ide_Z_Win"
Option Compare Binary
Option Explicit
Const CMod$ = "MVb_Ide_Z_Win."

Sub XAlign_WinV()
WinTileVBtn.Execute
End Sub

Function Ap_WinAy(ParamArray WinAp()) As VBIDE.Window()
Dim Av(): Av = WinAp
Dim I
For Each I In Av
    PushObj Ap_WinAy, I
Next
End Function

Property Get CdWinAy() As VBIDE.Window()
CdWinAy = WinTyWinAy(vbext_wt_CodeWindow)
End Property

Sub XClr_ImmWin()
DoEvents
With ImmWin
    .SetFocus
    .Visible = True
End With
Keys_XSnd "^{HOME}^+{END}"
DoEvents
'SndKeys "{DEL}" '<-- it does not work?
'DoEvents
End Sub

Sub ClsAllWin()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    W.Close
Next
End Sub

Sub XCls_AllWin_Exl(ParamArray ExlWinAp())
Dim W As VBIDE.Window, Exl()
Exl = ExlWinAp
For Each W In CurVbe.Windows
    If Not OyHas(Exl, W) Then
        W.Visible = False
    End If
Next
Dim I
For Each I In Exl
    CvWin(I).Visible = True
Next
End Sub

Sub XCls_CdWin_Exl(ExlMdNm$)
XCls_AllWin_Exl Md_Win(Md(ExlMdNm))
End Sub

Sub XCls_ImmWin()
DoEvents
ImmWin.Visible = False
End Sub

Sub XCls_Win_Exl(A As vbext_WindowType)
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    If W.Type <> A Then
        W.Close
    End If
Next
End Sub

Sub XCls_Win_ExlImm()
XCls_Win_Exl VBIDE.vbext_wt_Immediate
End Sub

Property Get CurCdWin() As VBIDE.Window
Set CurCdWin = CurVbe.ActiveCodePane.Window
End Property

Private Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Property Get CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
End Property

Function CvWin(A) As VBIDE.Window
Set CvWin = A
End Function

Function CvWinAy(A) As VBIDE.Window()
CvWinAy = A
End Function

Property Get EmpWinAy() As VBIDE.Window()
End Property

Property Get ImmWin() As VBIDE.Window
Set ImmWin = WinTyWin(vbext_wt_Immediate)
End Property

Property Get LclWin() As VBIDE.Window
Set LclWin = WinTyWin(vbext_wt_Locals)
End Property

Private Function Md(MdDNm) As CodeModule
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Sz(A1)
Case 1: Set Md = Pj_Md(CurVbe.ActiveVBProject, MdDNm)
Case 2: Set Md = Pj_Md(CurVbe.VBProjects(A1(0)), A1(1))
Case Else: XThw CSub, "MdDNm should be XXX.XXX or XXX", "MdDNm", MdDNm
End Select
End Function

Function Md_Win(A As CodeModule) As VBIDE.Window
Set Md_Win = A.CodePane.Window
End Function

Private Function Pj_Md(A As VBProject, Nm) As CodeModule
Set Pj_Md = A.VBComponents(Nm).CodeModule
End Function

Sub XShw_Dbg()
XCls_AllWin_Exl ImmWin, LclWin, CurCdWin
DoEvents
XAlign_WinV
XClr_ImmWin
End Sub

Sub XShw_NxtStmt()
Const CSub$ = CMod & "XShw_NxtStmt"
With ShwNxtStmtBtn
    If Not .Enabled Then
        Msg CSub, "ShwNxtStmtBtn is disabled"
        Exit Sub
    End If
    .Execute
End With
End Sub

Property Get VisWinCnt&()
VisWinCnt = Itr_Cnt_PrpTrue(CurVbe.Windows, "Visible")
End Property

Sub Win_XClr(A As VBIDE.Window)
DoEvents
IdeSelAllBtn.Execute
DoEvents
SendKeys " "
'IdeClrBtn.Execute
End Sub

Sub Win_XCls(A As VBIDE.Window)
If IsNothing(A) Then Exit Sub
If WinTy(A) = -1 Then Exit Sub
If A.Visible Then A.Close
End Sub

Property Get NWin&()
NWin = CurVbe.Windows.Count
End Property

Function Win_MdNm$(A As VBIDE.Window)
Win_MdNm = TakBet(A.Caption, " - ", " (Code)")
End Function

Property Get WinNy() As String()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    Debug.Print W.Caption, W.Type
    PushI WinNy, W.Caption
Next
End Property

Function WinTy(A As VBIDE.Window) As VBIDE.vbext_WindowType
On Error GoTo X
WinTy = A.Type
Exit Function
X: WinTy = -1
End Function

Function WinTyWin(A As vbext_WindowType) As VBIDE.Window
Set WinTyWin = Itr_FstItm_PrpEqV(CurVbe.Windows, "Type", A)
End Function

Function WinTyWinAy(T As vbext_WindowType) As VBIDE.Window()
WinTyWinAy = Itr_XWh_PrpEqV(CurVbe.Windows, "Type", T)
End Function

Function Win_XVis(A As VBIDE.Window) As VBIDE.Window
A.Visible = True
A.WindowState = vbext_ws_Normal
Set Win_XVis = A
End Function

Private Sub Z_Md()
Debug.Print Md("MVb_Ide_Z_Win").Parent.Name
End Sub

Private Sub ZZ()
Dim A()
Dim B$
Dim C As vbext_WindowType
Dim D As Variant
Dim E As CodeModule
Dim F As VBIDE.Window
XAlign_WinV
Ap_WinAy A
XCls_AllWin_Exl A
XCls_CdWin_Exl B
XCls_ImmWin
XCls_Win_Exl C
XCls_Win_ExlImm
CvWin D
CvWinAy D
Md_Win E
XShw_Dbg
XShw_NxtStmt
Win_XClr F
Win_XCls F
WinTy F
WinTyWin C
WinTyWinAy C
Win_XVis F
End Sub

Private Sub Z()
Z_Md
End Sub
