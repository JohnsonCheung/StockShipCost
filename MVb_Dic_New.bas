Attribute VB_Name = "MVb_Dic_New"
Option Compare Binary
Option Explicit
Function New_SyDic(TermSslAy$()) As Dictionary
Dim L, T$, Ssl$
Dim O As New Dictionary
For Each L In AyNz(TermSslAy)
    Lin_TRstAsg L, T, Ssl
    If O.Exists(T) Then
        O(T) = Ay_XAdd_(O(T), Ssl_Sy(Ssl))
    Else
        O.Add T, Ssl_Sy(Ssl)
    End If
Next
Set New_SyDic = O
End Function

Function NewDicSSL(Ssl) As Dictionary
Dim O As New Dictionary, I
For Each I In AyNz(Ssl_Sy(Ssl))
    DicSetKv O, I, Empty
Next
Set NewDicSSL = O
End Function

Sub DicSetKv(O As Dictionary, K, V)
If O.Exists(K) Then
    Asg V, O(K)
Else
    O.Add K, V
End If
End Sub
Function New_Dic_LY(Ly$(), Optional JnSep$ = vbCrLf) As Dictionary
Dim O As New Dictionary
Dim I, T$, Rst$
For Each I In AyNz(Ly)
    Lin_TRstAsg I, T, Rst
    If T <> "" Then
        If O.Exists(T) Then
            O(T) = O(T) & JnSep & Rst
        Else
            O.Add T, Rst
        End If
    End If
Next
Set New_Dic_LY = O
End Function

Function NewDicVBL(A$, Optional JnSep$ = vbCrLf) As Dictionary
Set NewDicVBL = New_Dic_LY(SplitVBar(A), JnSep)
End Function

