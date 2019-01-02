Attribute VB_Name = "MApp_Exp"
Option Compare Binary
Option Explicit

Sub XExp_App()
Pth_XClr SrcPth
'SpecExp
'AppExpMd
'AppExpFrm
XExp_AppStru
End Sub
Sub XBrw_AppStru()
XExp_AppStru
Ft_XBrw AppStruFfn
End Sub
Sub XExp_AppStru()
Str_XWrt Stru, AppStruFfn
End Sub
Property Get AppStruFfn$()
AppStruFfn = SrcPth & "Stru.txt"

End Property
