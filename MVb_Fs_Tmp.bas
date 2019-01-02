Attribute VB_Name = "MVb_Fs_Tmp"
Option Compare Binary
Option Explicit

Function TmpCmd$(Optional Fdr$, Optional Fnn$)
TmpCmd = TmpFfn(".cmd", Fdr, Fnn)
End Function

Function TmpFb$(Optional Fdr$, Optional Fnn$)
TmpFb = TmpFfn(".accdb", Fdr, Fnn)
End Function

Function TmpFcsv$(Optional Fdr$, Optional Fnn$)
TmpFcsv = TmpFfn(".csv", Fdr, Fnn)
End Function

Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
Fnn = IIf(Fnn0 = "", TmpNm, Fnn0)
TmpFfn = TmpFdrPth(Fdr) & Fnn & Ext
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function

Function TmpFxm$(Optional Fdr$, Optional Fnn0$)
TmpFxm = TmpFfn(".xlsm", Fdr, Fnn0)
End Function
Property Get TmpRoot$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpRoot = X
End Property

Property Get TmpHom$()
Static X$
If X = "" Then X = Pth_XEns(TmpRoot & "App")
TmpHom = X
End Property

Sub TmpHomBrw()
Pth_XBrw TmpHom
End Sub

Property Get TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Property

Function TmpFdrPth$(Fdr$)
Dim A$
If Fdr <> "" Then A = Fdr & "\"
TmpFdrPth = Pth_XEns(TmpHom & A)
End Function

Property Get TmpPth$()
TmpPth = Pth_XEns(TmpHom & TmpNm & "\")
End Property

Sub TmpPth_XBrw()
Pth_XBrw TmpPth
End Sub

Property Get TmpPth_Fix$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPth_Fix = X
End Property

Property Get TmpPthHom$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthHom = X
End Property
