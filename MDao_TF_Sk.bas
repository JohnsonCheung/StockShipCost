Attribute VB_Name = "MDao_TF_Sk"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_TF_Sk."

Private Function BExpr$(A As Database, T, Sk$(), K_Itm_or_Ay)
Const CSub$ = CMod & "BExpr"
Dim K()
    K() = CvAyITM(K_Itm_or_Ay)
If Sz(Sk) <> Sz(K) Then
    XThw CSub, "Given K() sz does not match Dbt Sk sz", "Db T Sk-Sz Given-K-Sz Sk K", Db_Nm(A), T, Sz(Sk), Sz(K), Sk, K
End If

Dim U%, S$
U = UB(K)
Dim O$(), J%, V
For J = 0 To U
    If IsNull(K(J)) Then
        Push O, Sk(J) & " is null"
    Else
        V = K(J): GoSub GetS
        Push O, Sk(J) & "=" & S
    End If
Next
BExpr = Join(O, " and ")
Exit Function
GetS:
Select Case True
Case IsStr(V): S = "'" & V & "'"
Case IsDate(V): S = "#" & V & "#"
Case IsBool(V): S = IIf(V, "TRUE", "FALSE")
Case IsNumeric(V): S = V
Case Else: Stop
End Select
Return
End Function

Property Get Dbtfk_Val(A As Database, T, F, K)
Dim Rs As DAO.Recordset
    Dim Sk$(): Sk = Dbt_SkFny(A, T)
    Dim W$: W = BExpr(A, T, Sk, K)
    Q = QQ_Fmt("Select [?] from [?] where ?", F, T, W)
    Set Rs = A.OpenRecordset(Q)
Dbtfk_Val = Nz(Rs_Val(Rs), Empty)
End Property

Property Let Dbtfk_Val(A As Database, T, F, K, V)
Dim Rs As DAO.Recordset
    Dim Sk$(): Sk = Dbt_SkFny(A, T)
    Dim W$: W = BExpr(A, T, Sk, K)
    Q = QQ_Fmt("Select [?],? from [?] where ?", F, JnComma(Sk), T, W)
    Set Rs = A.OpenRecordset(Q)
If Rs_XHas_Rec(Rs) Then
    Dr_XIns_Rs ItmAddAy(V, CvAyITM(K)), Rs
Else
    Dr_XUpd_Rs Array(V), Rs
End If
End Property

Function TfkV(T, F, K)
TfkV = Dbtfk_Val(CurDb, T, F, K)
End Function

Private Sub Z()
Exit Sub
'Dbtfk_Val
'TfkV
End Sub
