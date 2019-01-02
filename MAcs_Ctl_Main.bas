Attribute VB_Name = "MAcs_Ctl_Main"
Option Compare Binary
Option Explicit
Const CMod$ = "MAcs_Ctl_Main."

'Assume there is Application.Forms("Main").Msg (TextBox)
'MMsg means Main.Msg (TextBox)
Sub XClr_MainMsg()
If WIsOk Then WTBox.Value = ""
End Sub

Sub Qnm_XSet_MainMsg(QryNm)
Msg_XSet_MainMsg "Running query: " & QryNm
End Sub

Sub Msg_XSet_MainMsg(A$)
If WIsOk Then TBoxSet WTBox, A
End Sub

Private Property Get WTBox() As Access.TextBox
Set WTBox = WFrm.Controls("Msg")
End Property

Private Property Get WFrm() As Access.Form
Set WFrm = Access.Forms("Main")
End Property

Private Property Get WIsOk() As Boolean
Const CSub$ = CMod & "WIsOk"
If Not Itr_XHas_Nm(Access.CurrentProject.AllForms, "Main") Then
    Msg CSub, "no Main form"
    Exit Property
End If
If Not Itr_XHas_Nm(Access.Forms, "Main") Then
    Msg CSub, "Main form is not open"
    Exit Property
End If
If Not Itr_XHas_Nm(Access.Forms("Main").Controls, "Msg") Then
    Msg CSub, "no `Msg` textbox in main form"
    Exit Property
End If
WIsOk = True
End Property

Private Sub Z_WIsOk()
Debug.Print WIsOk
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$
XClr_MainMsg
Qnm_XSet_MainMsg A
Msg_XSet_MainMsg B
End Sub

Private Sub Z()
Z_WIsOk
End Sub
