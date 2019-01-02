Attribute VB_Name = "MVb_Lg_Lgr"
Option Compare Binary
Option Explicit

Sub LgrBrw()
Ft_XBrw LgrFt
End Sub

Property Get LgrFilNo%()
LgrFilNo = Ft_Fno_ForApp(LgrFt)
End Property

Property Get LgrFt$()
LgrFt = LgrPth & "Log.txt"
End Property

Sub LgrLg(Msg$)
Dim F%: F = LgrFilNo
Print #F, NowStr & " " & Msg
If LgrFilNo = 0 Then Close #F
End Sub

Property Get LgrPth$()
Dim O$:
'O = WrkPth: Pth_XEns O
O = O & "Log\": Pth_XEns O
LgrPth = O
End Property
