Attribute VB_Name = "MDao_Lnk_Spec"
Option Compare Binary
Option Explicit

Function LSClnInpy(A) As String()
LSClnInpy = Ssl_Sy(RmvT1(Ay_FstT1(A, "A-Inp")))
End Function

Property Get LSLines$()
LSLines = Spnm_Lines("Lnk")
End Property
Private Function ActFldLy(ActInpy$(), LyFld$()) As String()
ActFldLy = Ay_XWh_T1InAy(LyFld, ActInpy)
End Function

Function ActInpy(FmIp$(), InAct$()) As String()
'Dim Inpy$():   Inpy = Ssl_Sy(Ay_XWh_RmvTT(NoT1, "Inp", "|")(0))
'ActInpy = AyMinus(Inpy, InAct)
End Function


Sub LSpecAsg(A, Optional OTbl_Nm$, Optional OLnkColVbl$, Optional OWhBExpr$)
Dim Ay$()
Ay = AyTrim(SplitVBar(A))
OTbl_Nm = Ay_XShf_(Ay)
If Lin_T1(Ay_LasEle(Ay)) = "Where" Then
    OWhBExpr = RmvT1(Pop(Ay))
Else
    OWhBExpr = ""
End If
OLnkColVbl = JnVBar(Ay)
End Sub

Sub LSpecDmp(A)
Debug.Print XRpl_VBar(A)
End Sub

Function LSpecLy(A) As String()
Const L2Spec$ = ">GLAnp |" & _
    "Whs    Txt Plant |" & _
    "Loc    Txt [Storage Location]|" & _
    "Sku    Txt Material |" & _
    "PstDte Txt [Posting Date] |" & _
    "MovTy  Txt [Movement Type]|" & _
    "Qty    Txt Quantity|" & _
    "BchNo  Txt Batch |" & _
    "Where Plant='8601' and [Storage Location]='0002' and [Movement Type] like '6*'"
End Function

