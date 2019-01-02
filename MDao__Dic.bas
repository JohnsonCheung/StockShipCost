Attribute VB_Name = "MDao__Dic"
Option Compare Binary
Option Explicit
Function TF_ASet(T, F) As ASet
Set TF_ASet = Dbtf_ASet(CurDb, T, F)
End Function

Function Dbtf_ASet(A As Database, T, F) As ASet
Set Dbtf_ASet = Rs_ASet(Dbq_Rs(A, QSelDis_FF_Fm(F, T)))
End Function

Function Dbtf_Dic_CNT(A As Database, T, F) As Dictionary
Set Dbtf_Dic_CNT = Rs_Dic_CNT_COL1(Dbq_Rs(A, QSelDis_FF_Fm(F, T)))
End Function

Function Dbq_Dic_SY(A As Database, Q) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Sy
Set Dbq_Dic_SY = Rs_Dic_SY(Dbq_Rs(A, Q))
End Function
Function Dbq_AyDic(A As Database, Q) As Dictionary
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Ay
Set Dbq_AyDic = Rs_AyDic(Dbq_Rs(A, Q))
End Function
Function Rs_AyDic(A As DAO.Recordset) As Dictionary
Set Rs_AyDic = Rs_AyDic_INTO(A, EmpAy)
End Function
Function Rs_Dic_SY(A As DAO.Recordset) As Dictionary
Set Rs_Dic_SY = Rs_AyDic_INTO(A, EmpSy)
End Function

Function Rs_AyDic_INTO(A As DAO.Recordset, OInto) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Ay
Dim O As New Dictionary, K, V, Ay, ClnAy
ClnAy = Ay_XCln(OInto)
With A
    While Not .EOF
        K = .Fields(0).Value
        V = .Fields(1).Value
        If O.Exists(K) Then
            If True Then
                PushI O(K), V
            Else
                Ay = O(K)
                PushI Ay, V
                O(K) = Ay
            End If
        Else
            O.Add K, Array(0)
        End If
        .MoveNext
    Wend
End With
Set Rs_AyDic_INTO = O
End Function

Function Dbt_ASet_SK(A As Database, T$) As ASet
Set Dbt_ASet_SK = Dbtf_ASet(A, T, Dbt_SskFldNm(A, T))
End Function
Function Rs_Dic_OneRec(A As DAO.Recordset) As Dictionary
Dim F As DAO.Field
Dim O As New Dictionary
For Each F In A.Fields
    O.Add F.Name, F.Value
Next
Set Rs_Dic_OneRec = O
End Function
Function Rs_Dic_COL12(A As DAO.Recordset, Optional Sep$ = vbCrLf) As Dictionary _
'Return a Dic from col1 and col2 of Rs-A _
'Dic-Key: is distinct value of col1
'Dic-Val: is added-by-sep string val of col2
Dim O As New Dictionary
Dim K, V$
While Not A.EOF
    K = A.Fields(0).Value
    V = A.Fields(1).Value
    If O.Exists(K) Then
        O(K) = O(K) & Sep & V
    Else
        O.Add K, CStr(Nz(V))
    End If
    A.MoveNext
Wend
Set Rs_Dic_COL12 = O
End Function

Function Dbq_Dic(A As Database, Q, Optional Sep$ = vbCrLf & vbCrLf) As Dictionary
Set Dbq_Dic = Rs_Dic_COL12(Dbq_Rs(A, Q), Sep)
End Function

Function Rs_ASet_F(A As DAO.Recordset, F) As ASet
Dim O As ASet
Set O = New_ASet
With A
    While Not .EOF
        ASet_XPush O, .Fields(F).Value
        .MoveNext
    Wend
End With
Set Rs_ASet_F = O
End Function

Function Rs_ASet_COL1(A As DAO.Recordset) As ASet
Set Rs_ASet_COL1 = Rs_ASet_F(A, 0)
End Function
Function Rs_ASet(A As DAO.Recordset) As ASet
Set Rs_ASet = Rs_ASet_COL1(A)
End Function
Function Rs_Dic_CNT_COL1(A As DAO.Recordset) As Dictionary
Dim O As New Dictionary, V
While Not A.EOF
    V = A.Fields(0).Value
    If Not O.Exists(V) Then
        O.Add V, O.Count
    End If
    A.MoveNext
Wend
Set Rs_Dic_CNT_COL1 = O
End Function

