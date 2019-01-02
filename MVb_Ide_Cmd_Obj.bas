Attribute VB_Name = "MVb_Ide_Cmd_Obj"
Option Compare Binary
Option Explicit

Sub Bar_ClrAllCtl(A As CommandBar)
Dim I
For Each I In AyNz(Bar_CtlAy(A))
    CvCtl(I).Delete
Next
End Sub

Function Bar_CtlAy(A As CommandBar) As CommandBarControl()
Bar_CtlAy = Itr_Into(A.Controls, Bar_CtlAy)
End Function

Function Bar_CtlNy(A As CommandBar) As String()
End Function

Property Get BarNy() As String()
BarNy = Vbe_BarNy(CurVbe)
End Property

Property Get BrwObjWin() As VBIDE.Window
Set BrwObjWin = WinTyWin(vbext_wt_Browser)
End Property

Property Get CompileBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = DbgPop.CommandBar.Controls(1)
If Not XHas_Pfx(O.Caption, "Compi&le") Then Stop
Set CompileBtn = O
End Property

Private Function CvCtl(A) As CommandBarControl
Set CvCtl = A
End Function

Property Get DbgPop() As CommandBarPopup
Set DbgPop = MnuBar.Controls("Debug")
End Property

Property Get IdeClrBtn() As Office.CommandBarButton
Set IdeClrBtn = Itr_FstItm_PrpEqV(PopEdt.Controls, "Caption", "C&lear")
End Property

Property Get IdeMnuBar() As Office.CommandBar
Set IdeMnuBar = CurVbe.CommandBars("Menu Bar")
End Property

Property Get IdeSelAllBtn() As Office.CommandBarButton
Set IdeSelAllBtn = Itr_FstItm_PrpEqV(PopEdt.Controls, "Caption", "Select &All")
End Property

Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function

Private Property Get MnuBar() As CommandBar
Set MnuBar = VbeMnuBar(CurVbe)
End Property

Property Get NxtStmtBtn() As CommandBarButton
Set NxtStmtBtn = DbgPop.Controls("Show Next Statement")
End Property

Private Property Get PopEdt() As Office.CommandBarPopup
Set PopEdt = Itr_FstItm_PrpEqV(IdeMnuBar.Controls, "Caption", "&Edit")
End Property

Property Get SavBtn() As CommandBarButton
Set SavBtn = Vbe_SavBtn(CurVbe)
End Property

Property Get ShwNxtStmtBtn() As CommandBarButton
Set ShwNxtStmtBtn = DbgPop.Controls("Show Next Statement")
End Property

Property Get StdBar() As Office.CommandBar
Set StdBar = CurVbe_CmdBars("Standard")
End Property

Function VbeMnuBar(A As Vbe) As CommandBar
Set VbeMnuBar = A.CommandBars("Menu Bar")
End Function

Function Vbe_SavBtn(A As Vbe) As CommandBarButton
Dim I As CommandBarControl, S As Office.CommandBarControls
Set S = VbeStdBar(A).Controls
For Each I In S
    If XHas_Pfx(I.Caption, "&Sav") Then Set Vbe_SavBtn = I: Exit Function
Next
Stop
End Function

Function VbeStdBar(A As Vbe) As Office.CommandBar
Dim X As Office.CommandBars
Set X = Vbe_CmdBars(A)
Set VbeStdBar = X("Standard")
End Function

Function Vbe_BarNy(A As Vbe) As String()
Vbe_BarNy = Vbe_CmdBarNy(A)
End Function

Function Vbe_CmdBarAy(A As Vbe) As Office.CommandBar()
Dim I
For Each I In A.CommandBars
   PushObj Vbe_CmdBarAy, I
Next
End Function

Function Vbe_CmdBarNy(A As Vbe) As String()
Vbe_CmdBarNy = Itr_Ny(A.CommandBars)
End Function

Property Get WinPop() As CommandBarPopup
Set WinPop = MnuBar.Controls("Window")
End Property

Property Get WinTileVBtn() As Office.CommandBarButton
Set WinTileVBtn = WinPop.Controls("Tile &Vertically")
End Property

Property Get XlsBtn() As Office.CommandBarControl
Set XlsBtn = StdBar.Controls(1)
End Property

Private Sub ZZ_DbgPop()
Dim A
Set A = DbgPop
Stop
End Sub

Private Sub ZZ_MnuBar()
Dim A As CommandBar
Set A = MnuBar
Stop
End Sub

Private Sub ZZ()
Dim A As CommandBar
Dim B As Variant
Dim C As Vbe
Bar_ClrAllCtl A
Bar_CtlAy A
Bar_CtlNy A
IsBtn B
VbeMnuBar C
Vbe_SavBtn C
VbeStdBar C
Vbe_BarNy C
Vbe_CmdBarAy C
Vbe_CmdBarNy C
End Sub

Private Sub Z()
End Sub
