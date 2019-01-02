Attribute VB_Name = "MVb_Lin"
Option Compare Binary
Option Explicit
Function LinCnt&(Lines)
LinCnt = SubStrCnt(Lines, vbCrLf) + 1
End Function

Function LinHasDDRmk(A$) As Boolean
LinHasDDRmk = HasSubStr(A, "--")
End Function

Function LinHasLikItm(A, Lik$, Itm$) As Boolean
Dim L$, I$
AyAsg Lin_TT(A), L, I
If Not Lik Like L Then Exit Function
LinHasLikItm = I = Itm
End Function

Function Lin_IsSngTerm(A) As Boolean
Lin_IsSngTerm = InStr(Trim(A), " ") = 0
End Function

Function Lin_IsDDLin(A) As Boolean
Lin_IsDDLin = XTak_FstTwoChr(LTrim(A)) = "--"
End Function

Function Lin_IsDotLin(A) As Boolean
Lin_IsDotLin = XTak_FstChr(A) = "."
End Function

Function Lin_IsInT1Ay(A, T1Ay$()) As Boolean
Lin_IsInT1Ay = Ay_XHas(T1Ay, Lin_T1(A))
End Function

Function LinPfx$(A, ParamArray PfxAp())
Dim Av(): Av = PfxAp
Dim X
For Each X In Av
    If XHas_Pfx(A, X) Then LinPfx = X: Exit Function
Next
End Function

Function LinPfxErMsg$(Lin, Pfx$)
If XHas_Pfx(Lin, Pfx) Then Exit Function
LinPfxErMsg = QQ_Fmt("First Char must be [?]", Pfx)
End Function
