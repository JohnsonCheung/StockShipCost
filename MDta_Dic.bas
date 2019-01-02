Attribute VB_Name = "MDta_Dic"
Option Compare Binary
Option Explicit
Function Dic_Drs(A As Dictionary, Optional InclDicValTy As Boolean, Optional Tit$ = "Key Val") As Drs
Dim Fny$()
Fny = Ssl_Sy(Tit): If InclDicValTy Then Push Fny, "Val-TypeName"
Set Dic_Drs = New_Drs(Fny, DicDry(A, InclDicValTy))
End Function

Function NewDic_BY_DIC_DRY(DicDry()) As Dictionary
Dim O As New Dictionary
If Sz(DicDry) > 0 Then
   Dim Dr
   For Each Dr In DicDry
       O.Add Dr(0), Dr(1)
   Next
End If
Set NewDic_BY_DIC_DRY = O
End Function

Function Dic_Dt(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValTy As Boolean) As Dt
Dim Dry()
Dry = DicDry(A, InclDicValTy)
Dim F$
    If InclDicValTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
Set Dic_Dt = New_Dt(DtNm, F, Dry)
End Function

Function Dic_Fny(Optional InclValTy As Boolean) As String()
Dic_Fny = Ssl_Sy("Key Val" & IIf(InclValTy, " Type", ""))
End Function
