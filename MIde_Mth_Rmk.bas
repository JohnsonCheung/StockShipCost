Attribute VB_Name = "MIde_Mth_Rmk"
Option Compare Binary
Option Explicit

Sub XRmk_Mth()
Mth_XRmk CurMth
End Sub

Sub XUnRmk_Mth()
Mth_XUnRmk CurMth
End Sub

Sub Mth_XUnRmk(A As Mth)
Dim P() As FTNo: P = Mth_MthCxtFTNoAy(A)
Dim J%, FTNo As FTNo
For J = UB(P) To 0 Step -1
    Set FTNo = FTNo_XNxt_FmNo_IfItIsStopLin(P(J), A.Md.Lines(P(J).FmNo, 1))
    MdFTNo_XUmRmk A.Md, FTNo
    MdLno_XRmv_StopLin_IfAny A.Md, P(J).FmNo
Next
End Sub
Function FTNo_XNxt_FmNo(A As FTNo) As FTNo
Set FTNo_XNxt_FmNo = New_FTNo(A.FmNo + 1, A.ToNo)
End Function

Private Function FTNo_XNxt_FmNo_IfItIsStopLin(A As FTNo, Lin$) As FTNo
If Lin = "Stop '" Then
    Set FTNo_XNxt_FmNo_IfItIsStopLin = FTNo_XNxt_FmNo(A)
Else
    Set FTNo_XNxt_FmNo_IfItIsStopLin = A
End If
End Function
Private Sub ZZ_Mth_XRmk()
Dim M As Mth: Set M = New_Mth(Md("ZZModule"), "YYA")
              Ass Lines_Vbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
Mth_XRmk M:   Ass Lines_Vbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
Mth_XUnRmk M: Ass Lines_Vbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub

Sub Mth_XRmk(A As Mth)
Dim P() As FTNo: P = Mth_MthCxtFTNoAy(A)
D FTNoAy_Ly(P)
Dim J%
For J = UB(P) To 0 Step -1
    MdFTNo_XRmk A.Md, P(J)
    MdLno_XIns_StopLin_IfNeed A.Md, P(J).FmNo
Next
End Sub
