Attribute VB_Name = "MVb__Nm"
Option Compare Binary
Option Explicit

Function IsNm(A) As Boolean
If Not IsLetter(XTak_FstChr(A)) Then Exit Function
Dim L%: L = Len(A)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(A$) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
End Function

Function Nm_IsSel(A$, B As WhNm) As Boolean
If IsNothing(B) Then Nm_IsSel = True: Exit Function
Nm_IsSel = Nm_IsSel_ByReExl(A, B.Re, B.ExlAy)
End Function

Function Nm_IsSel_ByExlLikAy(A$, ExlLikAy$()) As Boolean
Nm_IsSel_ByExlLikAy = Not IsInLikAy(A, ExlLikAy)
End Function

Function Nm_IsSel_ByRe(A$, Re As RegExp) As Boolean
If A = "" Then Exit Function
If IsNothing(Re) Then Nm_IsSel_ByRe = True: Exit Function
Nm_IsSel_ByRe = Re.Test(A)
End Function

Function Nm_IsSel_ByReExl(A$, Re As RegExp, ExlLikAy$()) As Boolean
If Not Nm_IsSel_ByRe(A, Re) Then Exit Function
If Not Nm_IsSel_ByExlLikAy(A, ExlLikAy) Then Exit Function
Nm_IsSel_ByReExl = True
End Function

Function NmSfx$(A)
Dim J%, O$, C$
For J = Len(A) To 1 Step -1
    C = Mid(A, J, 1)
    If Not Asc_IsUCase(Asc(C)) Then
        If C <> "_" Then
            NmSfx = O: Exit Function
        End If
    End If
    O = C & O
Next
End Function
