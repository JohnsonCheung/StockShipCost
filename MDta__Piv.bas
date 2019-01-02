Attribute VB_Name = "MDta__Piv"
Option Explicit
Public Enum eAgg
    eSum
    eCnt
    eAvg
End Enum
Function DryGpAy(A, KIx%, GIx%) As Variant()
If Sz(A) = 0 Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
    K = Dr(KIx)
    Gp = Dr(GIx)
    O_Ix = Ay_Ix(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryGpAy = O
End Function

Function Aydic_to_KeyCntMulItmColDry(A As Dictionary) As Variant()
If A.Count = 0 Then Exit Function
Dim O(), K, Dr(), Ay, J&
ReDim O(A.Count - 1)
For Each K In A.Keys
    Ay = A(K): If Not IsArray(Ay) Then Stop
    O(J) = AyIns2(Ay, K, Sz(Ay))
    J = J + 1
Next
Aydic_to_KeyCntMulItmColDry = O
End Function
Function Drs_GpDic(A As Drs, KK, G$) As Dictionary
Dim Fny$()
Dim KeyIxAy&(), GIx%
    Fny = FF_Fny(KK)
    KeyIxAy = Ay_IxAy(A.Fny, Fny)
    PushI Fny, G & "_Gp"
    GIx = Ay_Ix(Fny, G)
Set Drs_GpDic = Dry_GpDic(A.Dry, KeyIxAy, GIx)
End Function
Function Dry_GpDic(A, KeyIxAy, G) As Dictionary
'If K < 0 Or G < 0 Then
'    XThw CSub, "K-Idx and G-Idx should both >= 0", "K-Idx G-Idx", K, G
'End If
'Dim Dr, U&, O As New Dictionary, KK, GG, Ay()
'U = UB(A): If U = -1 Then Exit Function
'For Each Dr In A
'    KK = Dr(K)
'    GG = Dr(G)
'    If O.Exists(KK) Then
'        Ay = O(KK)
'        PushI Ay, GG
'        O(KK) = Ay
'    Else
'        O.Add KK, Array(GG)
'    End If
'Next
'Set Dry_GpDic = O
End Function

Function KE24Drs() As Drs
Set KE24Drs = Dbt_Drs(Samp_Db_Duty_Dta, "KE24")
End Function

Function Dry_PivDry(A, KeyIxAy, G, Optional Agg As eAgg = eAgg.eSum) As Variant()
Dry_PivDry = Aydic_to_KeyCntMulItmColDry(Dry_GpDic(A, KeyIxAy, G))
End Function

Function Drs_PivDrs(A As Drs, KK, G$, Optional Agg As eAgg = eAgg.eSum) As Drs
Dim Dry(), S$, N%, KIxAy&(), GIx&
GIx = Ay_Ix(A.Fny, G)
KIxAy = Ay_IxAy(A.Fny, FF_Fny(KK))
Dry = Dry_PivDry(A.Dry, KIxAy, GIx, Agg)
N = Dry_NCol(Dry) - 2
S = LblSeqSsl(G, N)
'Fny0 = QQ_Fmt("? N ?", K, S)
Stop '
'Set Drs_PivDrs = New_Drs(Fny, Dry)
End Function

