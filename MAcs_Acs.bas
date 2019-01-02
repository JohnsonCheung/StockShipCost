Attribute VB_Name = "MAcs_Acs"
Option Compare Binary
Option Explicit
Sub Acs_SavRec(A As Access.Application)
A.DoCmd.RunCommand acCmdSaveRecord
End Sub

Sub AcsFb_XOpn(A As Access.Application, Fb$)
Select Case True
Case IsNothing(A.CurrentDb)
    A.OpenCurrentDatabase Fb
Case A.CurrentDb.Name = Fb
Case Else
    A.CurrentDb.Close
    A.OpenCurrentDatabase Fb
End Select
End Sub

Function Fb_XOpn(A$) As Access.Application
Acs.CloseCurrentDatabase
Acs.OpenCurrentDatabase A
Set Fb_XOpn = Acs
End Function

Sub Acs_XQuit(A As Access.Application)
Acs_XClsDb A
A.Quit
Set A = Nothing
End Sub

Sub Acs_XVis(A As Access.Application)
If Not A.Visible Then A.Visible = True
End Sub

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function
Function Fb_Acs(A, Optional Vis As Boolean) As Access.Application
Dim O As New Access.Application
O.OpenCurrentDatabase A
O.Visible = Vis
Set Fb_Acs = O
End Function

Function New_Acs(Optional Hid As Boolean) As Access.Application
Dim O As New Access.Application
If Not Hid Then O.Visible = True
Set New_Acs = O
End Function


Sub Acs_XCls(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Sub Acs_XClsDb(A As Access.Application)
Acs_XCls A
End Sub

Property Get Acs() As Access.Application
Static X As Boolean, Y As Access.Application
On Error GoTo X
If X Then
    Set Y = New Access.Application
    X = True
End If
If Y.Application.Name = "Microsoft Access" Then
    Set Acs = Y
    Exit Property
End If
X:
    Set Y = New Access.Application
    Debug.Print "Acs: New Acs instance is crreated."
Set Acs = Y
End Property

Sub TT_Cls(TT$)
Ay_XDo CvNy(TT), "Tbl_Cls"
End Sub

Sub Tbl_Cls(T)
CurAcs.DoCmd.Close acTable, T
End Sub

