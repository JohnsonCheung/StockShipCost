Attribute VB_Name = "MXls_Dao_Fb"
Option Compare Binary
Option Explicit
Function Fb_Wb_OUP_TBL(A$) As Workbook
Dim O As Workbook
Set O = New_Wb
Ay_XDo_ABX Fb_Tny_OUP(A), "Wb_XAdd_Wc", O, A
Itr_XDo O.Connections, "Wc_XAdd_Ws"
Wb_XRfh O, A
Set Fb_Wb_OUP_TBL = O
End Function

Sub FbWb_XRpl_Lo(Fb$, A As Workbook)
Dim I, Lo As ListObject, Db As Database
Set Db = Fb_Db(Fb)
For Each I In Wb_OupLoAy(A)
    Dbt_XRpl_Lo Db, "@" & Mid(Lo.Name, 3), CvLo(I)
Next
Db.Close
Set Db = Nothing
End Sub

Sub Fb_Fx_OUP_TBL(A$, Fx$)
Fb_Wb_OUP_TBL(A).SaveAs Fx
End Sub
