Attribute VB_Name = "MIde__ContLin"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde__ContLin."
Function MdContLin$(A As CodeModule, Lno)
Dim J&, L&
L = Lno
Dim O$: O = A.Lines(L, 1)
While XTak_LasChr(O) = "_"
    L = L + 1
    O = XRmv_LasChr(O) & A.Lines(L, 1)
Wend
MdContLin = O
End Function

Private Sub Z_SrcIx_ContLin()
Dim Src$(), MthIx%
MthIx = 0
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Src = O
Ept = "ABC"
GoSub Tst
Exit Sub
Tst:
    Act = SrcIx_ContLin(Src, MthIx)
    C
    Return
End Sub
Function SrcIx_ContLin$(A$(), Ix)
Const CSub$ = CMod & "SrcIx_ContLin"
If Ix <= -1 Then Exit Function
Dim J&, I$
Dim O$, IsCont As Boolean
For J = Ix To UB(A)
   I = A(J)
   O = O & LTrim(I)
   IsCont = XHas_Sfx(I, " _")
   If IsCont Then O = RmvSfx(RmvSfx(O, "_"), " ")
   If Not IsCont Then Exit For
Next
If IsCont Then XThw_Msg CSub, "each lines {Src} ends with sfx _, which is impossible"
SrcIx_ContLin = O
End Function
