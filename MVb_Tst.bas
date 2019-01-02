Attribute VB_Name = "MVb_Tst"
Option Compare Binary
Option Explicit
Public Act, Ept, Dbg As Boolean, Trc As Boolean

Sub C()
IsEq_XAss Act, Ept
Exit Sub
If Not IsEq(Act, Ept) Then
    XShw_Dbg
    D "=========================="
    D "Act"
    D Act
    D "---------------"
    D "Ept"
    D Ept
    Stop
End If
End Sub

Function PjTstPth$(Pj_Nm$)
PjTstPth = Pth_XEns(TstHom & Pj_Nm & "\")
End Function

Sub PjTstPth_XBrw(Pj_Nm$)
Pth_XBrw PjTstPth(Pj_Nm)
End Sub

Private Function TstCasPth$(Pj_Nm$, CSub$, CAs$)
TstCasPth = PjTstPth(Pj_Nm) & Ap_JnPthSep(Replace(CSub, ".", PthSep), CAs) & PthSep
End Function

Property Get TstHom$()
Static X$
Dim P$
P = Ffn_Pth(Application.Vbe.ActiveVBProject.FileName)
If X = "" Then X = Pth_XEns(P & "TstRes\")
TstHom = X
End Property

Sub TstHomBrw()
Pth_XBrw TstHom
End Sub

Sub TstOk(CSub$, CAs$)
Debug.Print "Tst OK | "; CSub; " | Case "; CAs
End Sub

Function TstTxt$(Pj_Nm$, CSub$, CAs$, Itm$, Optional IsEdt As Boolean)
If IsEdt Then
    TstTxtEdt Pj_Nm, CSub, CAs, Itm
    Exit Function
End If
TstTxt = Ft_Lines(TstTxtFt(Pj_Nm, CSub, CAs, Itm))
End Function

Sub TstTxtEdt(Pj_Nm$, CSub$, CAs$, Itm$)
Ft_XBrw TstTxtFt(Pj_Nm, CSub, CAs, Itm)
End Sub

Private Function TstTxtFt$(Pj_Nm$, CSub$, CAs$, Itm$)
TstTxtFt = FfnXEnsPth(TstCasPth(Pj_Nm, CSub, CAs) & Itm & ".txt")
End Function
