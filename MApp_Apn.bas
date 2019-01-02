Attribute VB_Name = "MApp_Apn"
Option Compare Binary
Option Explicit
Property Get AppPth$()
AppPth = Pth_XEns(TmpHom & Apn)
End Property
Private Sub Z_Apn()
Fb_CurDb Samp_Fb_ShpRate
Debug.Print Apn
End Sub
Property Get Apn$()
Static Y$
If Y = "" Then Y = Sql_Val("Select Apn from [Apn]")
Apn = Y
End Property

Function Apn_WDb(A) As Database
Static X As Boolean, Y As Database
If Not X Then
    X = True
    Fb_XEns Apn_WFb(A)
    Set Y = Fb_Db(Apn_WFb(A))
End If
If Not Db_IsOk(Y) Then Set Y = Fb_Db(Apn_WFb(A))
Set Apn_WDb = Y
End Function

Function Apn_WFb$(A)
Apn_WFb = ApnWPth(A) & "Wrk.accdb"
End Function

Function ApnWPth$(A)
Dim P$
P = TmpHom & A & "\"
Pth_XEns P
ApnWPth = P
End Function

Property Get PgmObjPth$()
PgmObjPth = Pth_XEns(CurDb_Pth & "PgmObj\")
End Property

Private Property Get PgmPth$()
PgmPth = Ffn_Pth(excel.Application.Vbe.ActiveVBProject.FileName)
End Property



Private Sub Z()
Z_Apn
MApp_Apn:
End Sub
