Attribute VB_Name = "MDao_Lnk_LnkColVbl"
Option Compare Binary
Option Explicit

Function ColLnk_ImpSql$(A$(), Fm)
'data ColLnk = F T E
Dim Into$, Ny$(), Ey$()
If XTak_FstChr(Fm) <> ">" Then Stop
Into = "#I" & Mid(Fm, 2)
Ny = Ay_XTak_T1(A)
Ey = AyMap_Sy(A, "RmvTT")
ColLnk_ImpSql = QSel_FF_ExprAy_Into_Fm_OWh$(Ny, Ey, Into, Fm)
End Function

Function Lin_LnkCol(A) As LnkCol
Dim Nm$, TyStr$, ExtNm$, Ty$
Lin_2TRstAsg A, Nm, Ty, ExtNm
ExtNm = XRmv_SqBkt(Trim(ExtNm))
Set Lin_LnkCol = New_LnkCol(Nm, Ty, IIf(ExtNm = "", Nm, ExtNm))
End Function

Function LnkColAy_ExtNy(A() As LnkCol) As String()
LnkColAy_ExtNy = Oy_PrpSy(A, "Extnm")
End Function

Function LnkColAy_SemiColonDaoShtTyStrAy(A() As LnkCol) As String()
LnkColAy_SemiColonDaoShtTyStrAy = Oy_PrpSy(A, "SemiColonDaoShtTyStr")
End Function

Function LnkColAy_Ny(A() As LnkCol) As String()
LnkColAy_Ny = Oy_PrpSy(A, "Nm")
End Function

Function LnkColVbl_LnkColAy(A) As LnkCol()
Dim L
For Each L In AyNz(SplitVBar(A))
    PushObj LnkColVbl_LnkColAy, Lin_LnkCol(L)
Next
End Function

Sub LnkColVbl_NyExtNy_XAsg(A$, ONy$(), OExtNy$())
Dim Ay() As LnkCol
    Ay = LnkColVbl_LnkColAy(A)
ONy = LnkColAy_Ny(Ay)
OExtNy = LnkColAy_ExtNy(Ay)
End Sub

Function LnkColVbl_LnkColLy(A$) As String()
Dim A3$(), A1$(), A2$(), Ay() As LnkCol
Ay = LnkColVbl_LnkColAy(A)
A1 = LnkColAy_Ny(Ay)
A2 = LnkColAy_SemiColonDaoShtTyStrAy(Ay)
A3 = Ay_XQuote_SqBkt(LnkColAy_ExtNy(Ay))
Stop '
Dim J%, O$()
For J = 0 To UB(A1)
    Push O, A1(J) & "  " & A2(J) & " " & A3(J)
Next
LnkColVbl_LnkColLy = Ay_XAlign_2T(O)
End Function

Private Sub Z_Lin_LnkCol()
Dim A$, Act As LnkCol, Exp As LnkCol
A = "AA Txt;Dbl [XX XX]"
Set Exp = New_LnkCol("AA", "Txt;Dbl", "XX XX")
GoSub Tst
Exit Sub
Tst:
    Set Act = Lin_LnkCol(A)
    Debug.Assert LnkCol_IsEq(Act, Exp)
    Return
End Sub
Function New_LnkCol(Nm$, SemiColonDaoShtTyStr$, ExtNm$) As LnkCol
Dim O As New LnkCol
Set New_LnkCol = O.Init(Nm, SemiColonDaoShtTyStr, ExtNm)
End Function


Function LnkCol_IsEq(A As LnkCol, B As LnkCol) As Boolean
With A
    If .ExtNm <> B.ExtNm Then Exit Function
    If .SemiColonDaoShtTyStr <> B.SemiColonDaoShtTyStr Then Exit Function
    If .Nm <> B.Nm Then Exit Function
End With
LnkCol_IsEq = True
End Function

Private Sub Z()
Z_Lin_LnkCol
Exit Sub
'ColLnk_ImpSql
'Lin_LnkCol
'LnkColAy_ExtNy
'LnkColAy_Ny
'LnkColVbl_LnkColAy
'LnkColVbl_Ly
End Sub
