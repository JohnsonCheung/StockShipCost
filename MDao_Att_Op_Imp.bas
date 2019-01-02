Attribute VB_Name = "MDao_Att_Op_Imp"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_Att_Op_Imp."

Sub Att_XImp(A$, FmFfn$)
DbAtt_XImp CurDb, A, FmFfn
End Sub

Private Sub ZImp(A As AttRs, Ffn$)
Const CSub$ = CMod & "ZImp"
Dim F2 As Field2
Dim S&, T$
S = Ffn_Sz(Ffn)
T = FfnDTim(Ffn)
Msg CSub, "[Att] is going to import [Ffn] with [Sz] and [Tim]", Fd_Val(A.Tbl_Rs!AttNm), Ffn, S, T
With A
    .Tbl_Rs.Edit
    With .AttRs
        If Rs_XHas_FldEqV(A.AttRs, "FileName", Ffn_Fn(Ffn)) Then
            D "Ffn is found in Att and it is replaced"
            .Edit
        Else
            D "Ffn is not found in Att and it is imported"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .Tbl_Rs.Fields!FilTim = Ffn_Tim(Ffn)
    .Tbl_Rs.Fields!FilSz = Ffn_Sz(Ffn)
    .Tbl_Rs.Update
End With
End Sub

Sub DbAtt_XImp(A As Database, Att$, FmFfn$)
ZImp DbAtt_AttRs(A, Att), FmFfn
End Sub

Private Sub Z_Att_XImp()
Dim T$
T = TmpFt
Str_XWrt "sdfdf", T
Att_XImp "AA", T
Kill T
'T = TmpFt
'Att_XExp_ToFfn "AA", T
'Ft_XBrw T
End Sub

Private Sub Z()
Z_Att_XImp
MDao_Att_Op_Imp:
End Sub
