Attribute VB_Name = "MApp_Pm"
Option Compare Binary
Option Explicit

Property Get Pnm_OupPth$()
Pnm_OupPth = Pnm_Val("OupPth")
End Property

Function Pnm_Ffn$(A)
Pnm_Ffn = Pnm_Pth(A) & Pnm_Fn(A)
End Function

Function Pnm_Fn$(A)
Pnm_Fn = Pnm_Val(A & "Fn")
End Function

Function Pnm_Pth$(A)
Pnm_Pth = Pth_XEns_Sfx(Pnm_Val(A & "Pth"))
End Function

Property Get Pnm_Val$(A)
Pnm_Val = CurDb.TableDefs("Pm").OpenRecordset.Fields(A).Value
End Property

Property Let Pnm_Val(A, V$)
Stop
'Should not use
With CurDb.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(A).Value = V
    .Update
End With
End Property

Sub XBrw_TblPm()
Tbl_XOpn "Pm"
End Sub

Private Sub ZZ()
Dim A$
Dim B As Variant
Pnm_Ffn A
Pnm_Ffn B
Pnm_Fn B
Pnm_Pth B
End Sub

Private Sub Z()
End Sub
