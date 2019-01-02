Attribute VB_Name = "MVb_Er"
Option Compare Binary
Option Explicit
Sub XThw_Msg(Fun$, Msg$)
XThw_Msg Fun, Msg
End Sub
Sub XThw_PgmErMsg(Fun$, Msg$)
XThw Fun, "Program Error: " & Msg, ""
End Sub
'Calling this module functions will throw error
Sub XThw(Fun$, Msg$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
XThwAv Fun, Msg, Ny0, Av
End Sub

Sub XThwAv(Fun$, MsgVbl$, Ny0, Av())
Ay_XBrw FunMsgNyAv_Ly(Fun, MsgVbl, Ny0, Av)
XHalt
End Sub
Sub Ass(A As Boolean)
Debug.Assert A
End Sub
Sub XAss(A As Boolean, Fun$, Msg$, FF$, ParamArray Ap())
If Not A Then
    Dim Av(): Av = Ap
    XThwAv Fun, Msg, FF, Av
End If
End Sub
Sub XThw_Never(Fun$)
XThw Fun, "Program should not reach here", ""
End Sub


Sub ChkAss(Chk$())
If Sz(Chk) = 0 Then Exit Sub
Ay_XBrw Chk
Stop
End Sub

Sub XHalt()
Err.Raise -1, , "Please check messages opened in notepad"
End Sub

Sub XHalt_Never(Fun$)
XThw Fun, "Should never reach here", ""
End Sub

