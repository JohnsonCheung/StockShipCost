Attribute VB_Name = "MVb_Stop_Xls"
Option Compare Binary
Option Explicit
Declare Function GetCurrentProcessId& Lib "Kernel32.dll" ()
Declare Function GetProcessId& Lib "Kernel32.dll" (ProcessHandle&)
'Const Ps1Str$ = "function Get-ExcelProcessId { try { (Get-Process -Name Excel).Id } finally { @() } }" & vbCrLf & _
'"Stop-Process -Id (Get-ExcelProcessId)"
Const WPs1Str$ = "Stop-Process -Id{try{(Get-Process -Name Excel).Id}finally{@()}}.invoke()"

Sub XStop_Xls()
WXEns
WXRun
End Sub
Private Sub WXEns()
Static X As Boolean
If Not X Then X = True: WXEns_Ps1Ffn
End Sub
Private Sub WXDmp_Ps1Cxt()
D WPs1Cxt
End Sub
Private Property Get WPs1Cxt$()
WPs1Cxt = Ft_Lines(WPs1Ffn)
End Property

Private Sub WXEns_Ps1Ffn()
'Ffn_XDltIfExist Ps1Ffn
If Ffn_Exist(WPs1Ffn) Then Exit Sub
Str_XWrt WPs1Str, WPs1Ffn
End Sub
Private Property Get WPs1Ffn$()
Static O$
If O = "" Then O = TmpHom & "StopXls.ps1"
WPs1Ffn = O
End Property
Private Sub WXRun()
Dim A$
A = QQ_Fmt("Powershell ""?""", WPs1Ffn)
'Debug.Print A
Shell A, vbHide
End Sub
