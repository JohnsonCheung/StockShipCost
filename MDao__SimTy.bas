Attribute VB_Name = "MDao__SimTy"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao__SimTy."
Function SimTy_QuoteTp$(A As eSimTy)
Const CSub$ = CMod & "SimTy_QuoteTp"
Dim O$
Select Case A
Case eTxt: O = "'?'"
Case eNbr, eLgc: O = "?"
Case eDte: O = "#?#"
Case Else
   XThw CSub, "Given {eSimTy} should be [eTxt eNbr eDte eLgc]", A
End Select
SimTy_QuoteTp = O
End Function

Function SimTyAy_InsValTp$(SimTyAy() As eSimTy)
Dim U%
   U = UB(SimTyAy)
Dim Ay$()
   ReDim Ay(U)
Dim J%
For J = 0 To U
   Ay(J) = SimTy_QuoteTp(SimTyAy(J))
Next
SimTyAy_InsValTp = JnComma(Ay)
End Function

Function SimTyStr_SimTy(A) As eSimTy
Dim O As eSimTy
Select Case UCase(A)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
SimTyStr_SimTy = O
End Function

Function IsSimTyLin(A) As Boolean
Stop '
End Function

Function IsSimTyss(A) As Boolean
Dim I
For Each I In AyNz(Ssl_Sy(A))
    If Not IsSimTyStr(CStr(I)) Then Exit Function
Next
IsSimTyss = True
End Function

Function IsSimTyStr(A$) As Boolean
Select Case UCase(A)
Case "TXT", "NBR", "LGC", "DTE", "OTH": IsSimTyStr = True
End Select
End Function
