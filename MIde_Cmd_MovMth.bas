Attribute VB_Name = "MIde_Cmd_MovMth"
Option Compare Binary
Option Explicit
Const MovMthBarNm$ = "MovMth"
Const MovMthBtnNm$ = "MovMth"

Property Get CmdBarNy() As String()
CmdBarNy = Itr_Ny(CurVbe_CmdBars)
End Property

Private Sub Z_XMov_MthBar()
MsgBox MovMthBar.Name
End Sub

Function Vbe_CmdBars(A As Vbe) As Office.CommandBars
Set Vbe_CmdBars = A.CommandBars
End Function

Property Get CurVbe_CmdBars() As Office.CommandBars
Set CurVbe_CmdBars = Vbe_CmdBars(CurVbe)
End Property

Function CurVbe_CmdBarsHas(A) As Boolean
CurVbe_CmdBarsHas = Itr_XHas_Nm(CurVbe_CmdBars, A)
End Function
Function CmdBar(A) As Office.CommandBar
Set CmdBar = CurVbe_CmdBars(A)
End Function
Sub RmvCmdBar(A)
If CurVbe_CmdBarsHas(A) Then CmdBar(A).Delete
End Sub
Function CvCmdBtn(A) As Office.CommandBarButton
Set CvCmdBtn = A
End Function
Function CmdBar_XHas_Btn(A As Office.CommandBar, BtnCaption)
Dim C As Office.CommandBarControl
For Each C In A.Controls
    If C.Type = msoControlButton Then
        If CvCmdBtn(C).Caption = BtnCaption Then CmdBar_XHas_Btn = True: Exit Function
    End If
Next
End Function
Sub XEns_CmdBarBtn(CmdBarNm, BtnCaption)
XEns_CmdBar MovMthBarNm
If CmdBar_XHas_Btn(CmdBar(CmdBarNm), BtnCaption) Then Exit Sub
CmdBar(CmdBarNm).Controls.Add(msoControlButton).Caption = BtnCaption
End Sub
Sub XEns_CmdBar(A$)
If CurVbe_CmdBarsHas(A) Then Exit Sub
XAdd_CmdBar A
End Sub
Sub XAdd_CmdBar(A)
CurVbe_CmdBars.Add A
End Sub
Property Get MovMthBar() As Office.CommandBar
Set MovMthBar = CurVbe_CmdBars(MovMthBarNm)
End Property
Property Get MovMthBtn() As Office.CommandBarControl
Set MovMthBtn = MovMthBar.Controls(MovMthBtnNm)
End Property

Private Sub Z()
Z_XMov_MthBar
MIde_Cmd_XMov_Mth:
End Sub
