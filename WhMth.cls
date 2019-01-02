VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhMth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Binary
Option Explicit
Public Nm As WhNm
Dim X_InMdy$(), X_InShtKd$()

Property Get InShtKd() As String()
InShtKd = X_InShtKd
End Property

Property Let InShtKd(V$())
X_InShtKd = V
End Property

Property Get InMdy() As String()
InMdy = X_InMdy
End Property

Property Let InMdy(V$())
X_InMdy = V
End Property
