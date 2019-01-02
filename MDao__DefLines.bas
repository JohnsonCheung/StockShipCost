Attribute VB_Name = "MDao__DefLines"
Option Compare Binary
Option Explicit

Function Fd_Lin$(A As DAO.Field)
Fd_Lin = Fd_Str(A)
End Function

Function Fds_Ly(A As DAO.Fields) As String()
Dim F As DAO.Field
For Each F In A
    PushI Fds_Ly, Fd_Lin(F)
Next
End Function

Function Idx_Lin$(A As DAO.Index)
Dim X$, F$
With A
Idx_Lin = QQ_Fmt("Idx;?;?;?", .Name, X, F)
End With
End Function

Function Idxs_Ly(A As DAO.Indexes) As String()
Dim I As DAO.Index
For Each I In A
    PushI Idxs_Ly, Idx_Lin(I)
Next
End Function
