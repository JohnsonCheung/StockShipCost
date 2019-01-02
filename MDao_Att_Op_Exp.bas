Attribute VB_Name = "MDao_Att_Op_Exp"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Att_Op_Exp."

Function Att_XExp_ToFfn$(A$, ToFfn)
'Exporting the only file in Att & Return ToFfn
Att_XExp_ToFfn = DbAtt_XExp_ToFfn(CurDb, A, ToFfn)
End Function

Function AttFn_XExp_ToFfn$(A$, AttFn$, ToFfn)
AttFn_XExp_ToFfn = DbAttFn_XExp_ToFfn(CurDb, A, AttFn, ToFfn)
End Function

Function AttRs_XExp_ToFfn$(A As AttRs, ToFfn)
'Export the only File in {AttRs} {ToFfn}
Dim Fn$, Ext$, T$, F2 As DAO.Field2
With A.AttRs
    If Ffn_Ext(!FileName) <> Ffn_Ext(ToFfn) Then Stop
    Set F2 = !FileData
End With
F2.SaveToFile ToFfn
AttRs_XExp_ToFfn = ToFfn
End Function

Function DbAtt_XExp_ToFfn$(A As Database, Att, ToFfn)
'Exporting the first File in Att.
'If no or more than one file in att, error
'If any, export and return ToFfn
Const CSub$ = CMod & "DbAtt_XExp_ToFfn"
Dim N%
N = DbAtt_NFil(A, Att)
If N <> 1 Then
    XThw CSub, "AttNm in Db should have a filecount of 1.  Cannot export.", _
        "Att-FileCount AttNm Db ExpToFile", _
        N, Att, Db_Nm(A), ToFfn
End If
DbAtt_XExp_ToFfn = AttRs_XExp_ToFfn(DbAtt_AttRs(A, Att), ToFfn)
XDmp_Ly CSub, "Att is exported", "Att ToFfn FmDb", Att, ToFfn, Db_Nm(A)
End Function

Function DbAttFn_XExp_ToFfn$(A As Database, Att$, AttFn$, ToFfn)
Const CSub$ = CMod & "DbAttFn_XExp_ToFfn"
If Ffn_Ext(AttFn) <> Ffn_Ext(ToFfn) Then
    XThw CSub, "AttFn & ToFfn are dif extension." & _
        "To export an AttFn to ToFfn, their file extension should be same", _
        "AttFn-Ext ToFfn-Ext Db AttNm AttFn ToFfn", _
        Ffn_Ext(AttFn), Ffn_Ext(ToFfn), Db_Nm(A), Att, AttFn, ToFfn
End If
If Ffn_Exist(ToFfn) Then
    XThw CSub, "ToFfn exist, no over write", _
        "Db AttNm AttFn ToFfn", _
        Db_Nm(A), Att, AttFn, ToFfn
End If
Dim Fd2 As DAO.Field2
    Set Fd2 = WFd2(A, Att, AttFn$)

If IsNothing(Fd2) Then
    XThw CSub, _
        "In record of AttNm there is no given AttFn, but only Act-Att_FnAy", _
        "Db Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        Db_Nm(A), Att, AttFn, DbAtt_FnAy(A, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
DbAttFn_XExp_ToFfn = ToFfn
End Function
Private Function WFd2(A As Database, Att, AttFn) As DAO.Field2
With DbAtt_AttRs(A, Att)
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !FileName = AttFn Then
                Set WFd2 = !FileData
            End If
            .MoveNext
        Wend
    End With
End With
End Function
Private Sub ZZ_DbAtt_XExp_ToFfn()
Dim T$
T = TmpFx
DbAttFn_XExp_ToFfn CurDb, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert Ffn_Exist(T)
Kill T
End Sub

Private Sub Z()
End Sub

Private Sub ZZ()
Dim A$
Dim B
Dim C As AttRs
Dim D As Database
Dim XX
AttFn_XExp_ToFfn A, A, B
AttRs_XExp_ToFfn C, B
Att_XExp_ToFfn A, B
DbAttFn_XExp_ToFfn D, A, A, B
DbAtt_XExp_ToFfn D, B, B
End Sub
