VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit
Private Sub CmdEditPrm_Click()
Me.UOMFn.Enabled = True
Me.UomPth.Enabled = True
Me.ZHT1Fn.Enabled = True
Me.ZHT1Pth.Enabled = True
Me.MB52Pth.Enabled = True
End Sub

Private Sub CmdOpnUomPth_Click()
Pth_XBrw Me.UomPth.Value
End Sub
Private Sub CmdOpnMB52Pth_Click()
Pth_XBrw Me.MB52Pth.Value
End Sub

Sub MsgSet(A$)
TBoxSet Msg, A
End Sub

Sub MsgClr()
Me.Msg.Value = ""
DoEvents
End Sub

Private Sub CmdOpnZHT1Pth_Click()
Pth_XBrw Me.ZHT1Pth.Value
End Sub

Private Sub CmdQuit_Click()
SavRec
Quit
End Sub

Private Sub CmdRunReport_Click()
If Not Vdt Then Exit Sub
XRpt
End Sub

Private Property Get Vdt() As Boolean
SavRec
MsgClr
If Pth_Er(UomPth) Then Exit Property
If FnEr(UOMFn, UomPth) Then Exit Property
If Pth_Er(MB52Pth) Then Exit Property
If FnEr(MB52Fn, MB52Pth) Then Exit Property
If MB52FnFmtEr(MB52Fn) Then Exit Property
DoCmd.RunCommand acCmdSaveRecord
Vdt = True
End Property

Private Function MB52FnFmtEr(A As TextBox) As Boolean
If Left(A.Value, 5) <> "MB52 " Then
    Me.Msg.Value = "MB52 file name must begin with [MB52 ]"
    A.SetFocus
    MB52FnFmtEr = True
    Exit Function
End If
Dim B$
B = Mid(A.Value, 6, 10)
If Not YYYYMMDD_IsVdt(B) Then
    Me.Msg.Value = "MB52 file name must begin with [MB52 YYYY-MM-DD]"
    A.SetFocus
    MB52FnFmtEr = True
End If
End Function
Private Function FnEr(A As TextBox, P As TextBox) As Boolean
If Dir(Pth_XEns_Sfx(P.Value) & A.Value) = "" Then
    A.Enabled = True
    A.SetFocus
    Me.Msg.Value = "File not found"
    FnEr = True
End If
End Function
Private Function Pth_Er(A As TextBox) As Boolean
If Dir(A.Value, vbDirectory) = "" Then
    A.Enabled = True
    A.SetFocus
    Me.Msg.Value = "Path not found"
    Pth_Er = True
End If
End Function

Private Sub CmOpnMB52Pth_Click()
Pth_XBrw Me.MB52Pth.Value
End Sub

Private Sub Form_Load()
Me.Msg.Value = ""
If IsDev Then
    Me.MB52Fn.Value = "MB52 2018-07-30.xls"
Else
    Me.MB52Fn.Value = "MB52 " & Format(Now, "YYYY-MM-DD") & ".xls"
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
XCpy_TblPm_Fm_C_or_N IsDev
Me.RecordSource = "Pm"
SavRec
End Sub

