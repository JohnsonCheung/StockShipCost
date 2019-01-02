Attribute VB_Name = "MVb_Dic_LinesDic"
Option Compare Binary
Option Explicit
Function IsLinesDic(A As Dictionary) As Boolean
If Not Ay_IsAllStr(A.Keys) Then Exit Function
IsLinesDic = Ay_IsLinesAy(A.Items)
End Function

Sub LinesDic_XBrw(A As Dictionary)
Ay_XBrw LinesDic_Fmt(A)
End Sub

Function LinesDic_Fmt(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    PushIAy LinesDic_Fmt, LinesDic_Fmt1(K, A(K))
Next
End Function

Private Function LinesDic_Fmt1(K, Lines) As String()
Dim L
For Each L In AyNz(SplitCrLf(Lines))
    Push LinesDic_Fmt1, K & " " & L
Next
End Function

Function New_Dic_LINES(A$()) As Dictionary
Dim O As New Dictionary
    Dim L, T1$
    For Each L In AyNz(A)
        T1 = XShf_Term(L)
        If O.Exists(T1) Then
            O(T1) = O(T1) & vbCrLf & L
        Else
            O(T1) = L
        End If
    Next
Set New_Dic_LINES = O
End Function

