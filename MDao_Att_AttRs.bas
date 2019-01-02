Attribute VB_Name = "MDao_Att_AttRs"
Option Compare Binary
Option Explicit

Function New_AttRs(A) As AttRs
New_AttRs = DbAtt_AttRs(CurDb, A)
End Function

Function AttRs_AttNm$(A As AttRs)
AttRs_AttNm = A.Tbl_Rs!AttNm
End Function

Function AttRs_NFil%(A As AttRs)
AttRs_NFil = RsNRec(A.AttRs)
End Function

Function AttRs_FstFn$(A As AttRs)
With A.AttRs
    If .EOF Then
        If .BOF Then
            Msg CSub, "[AttNm] has no attachment files", AttRs_AttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttRs_FstFn = !FileName
End With
End Function

Function DbAtt_AttRs(A As Database, Att) As AttRs
With DbAtt_AttRs
    Set .Tbl_Rs = A.OpenRecordset(QQ_Fmt("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .Tbl_Rs.EOF Then
        A.Execute QQ_Fmt("Insert into Att (AttNm) values('?')", Att)
        Set .Tbl_Rs = A.OpenRecordset(QQ_Fmt("Select Att from Att where AttNm='?'", Att))
    End If
    Set .AttRs = .Tbl_Rs.Fields(0).Value
End With
End Function

Function Db_FstAttRs(A As Database) As AttRs
With Db_FstAttRs
    Set .Tbl_Rs = A.TableDefs("Att").OpenRecordset
    Set .AttRs = .Tbl_Rs.Fields("Att").Value
End With
End Function
