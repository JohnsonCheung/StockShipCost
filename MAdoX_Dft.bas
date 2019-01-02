Attribute VB_Name = "MAdoX_Dft"
Option Compare Binary
Option Explicit
Function DftWsNy(WsNy0, Fx$) As String()
Dim O$()
    O = CvSy(WsNy0)
If Sz(O) = 0 Then
    DftWsNy = Fx_WsNy(Fx)
Else
    DftWsNy = O
End If
End Function
Function DftTny(Tny0, Fb$) As String()
Dim O$()
    O = CvSy(Tny0)
If Sz(O) = 0 Then
    DftTny = Fb_Tny(Fb)
Else
    DftTny = O
End If
End Function
Function FxDftWsNy(A, WsNy0) As String()
Dim O$(): O = CvSy(WsNy0)
If Sz(O) = 0 Then
    FxDftWsNy = Fx_WsNy(A)
Else
    FxDftWsNy = O
End If
End Function


Function FxDftWsNm$(A, WsNm0$)
If WsNm0 = "" Then
    FxDftWsNm = Fx_FstWsNm(A)
    Exit Function
End If
FxDftWsNm = WsNm0
End Function
