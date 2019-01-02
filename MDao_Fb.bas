Attribute VB_Name = "MDao_Fb"
Option Compare Binary
Option Explicit

Function Fb_XCrt(A) As Database
Set Fb_XCrt = DAO.DBEngine.CreateDatabase(A, dbLangGeneral)
End Function
Private Sub Z_Fb_XBrw()
Fb_XBrw Samp_Fb_Duty_Dta
End Sub
Sub Fb_XBrw(A)
Static Acs As New Access.Application
Acs.OpenCurrentDatabase A
Acs.Visible = True
End Sub
Function Fb_Cn_DAO(A) As DAO.Connection
Set Fb_Cn_DAO = DBEngine.OpenConnection(A)
End Function

Function Fb_Db(A) As Database
Set Fb_Db = DAO.DBEngine.OpenDatabase(A)
End Function

Sub Fb_XEns(A)
If Not Ffn_Exist(A) Then Fb_XCrt A
End Sub

Function Fb_Tny_OUP(A) As String()
Fb_Tny_OUP = Ay_XWh_Pfx(Fb_Tny(A), "@")
End Function

Sub FbtStr_XAsg(FbtStr$, OFb$, OT$)
If FbtStr = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
BrkAsg FbtStr, "].[", OFb, OT
If XTak_FstChr(OFb) <> "[" Then Stop
If XTak_LasChr(OT) <> "]" Then Stop
OFb = XRmv_FstChr(OFb)
OT = XRmv_LasChr(OT)
End Sub

Sub FbtDrp(A$, T$)
Dbt_XDrp Fb_Db(A), T
End Sub
Private Sub ZZ_FbTbl_Exist()
Ass FbTbl_Exist(Samp_Fb_Duty_Dta, "SkuB")
End Sub


Private Sub ZZ_Fb_Tny_OUP()
Dim Fb$
D Fb_Tny_OUP(Fb)
End Sub

Private Sub ZZ_Fb_Tny()
Ay_XDmp Fb_Tny(Samp_Fb_Duty_Dta)
End Sub


Private Sub Z()
Z_Fb_XBrw
MDao_Fb:
End Sub
