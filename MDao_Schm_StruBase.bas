Attribute VB_Name = "MDao_Schm_StruBase"
Option Compare Binary
Option Explicit
Type EF
    E As Dictionary ' Lookup Ele by Fld: Key=E; Val=FldLikss
    F As Dictionary ' Lookup Fd by Ele: Key=Ele; Val=Fd
End Type
Type StruBase
    EF As EF    'Ele FldLik
    Des As Dictionary ' Key=[T].[F]; Val=Des; That means Key must has substr [.]
End Type

Function NewEF(E As Dictionary, Optional F As Dictionary) As EF
Set NewEF.E = E
Set NewEF.F = F
End Function

Property Get EmpEF() As EF
End Property

Private Function ZFDic(Ly$()) As Dictionary
'FDic: Key=Ele;Val=Fd
'EleDefLin: $Ele *Ty ?Req
Dim E, A$, B$
Set ZFDic = New Dictionary
For Each E In AyNz(Ly)
    Lin_TRstAsg E, A, B
    ZFDic.Add Lin_T1(A), EleStr_Fd(B)
Next
End Function
Property Get EmpStruBase() As StruBase
End Property
Function New_StruBaseEF(EF As EF) As StruBase
New_StruBaseEF.EF = EF
End Function

Function New_StruBase(SchmLy$()) As StruBase
With New_StruBase
    Set .EF.E = New_Dic_LY(Ay_XWh_RmvT1(SchmLy, "Ele"))
    Set .EF.F = ZFDic(Ay_XWh_RmvT1(SchmLy, "Fld"))
    Set .Des = New_Dic_LY(Ay_XWh_RmvT1(SchmLy, "Des"))
End With
End Function

Private Sub Z_New_StruBase()
GoSub Cas1
GoSub Cas2
Exit Sub
Dim Ly$(), EptB As StruBase, ActB As StruBase
Cas2:
    GoTo Tst
Cas1:
Erase Ly
PushI Ly, "Tbl A *Id | *Nm | *Dte AATy Loc Expr Rmk"
PushI Ly, "Tbl B *Id | AId *Nm | *Dte"
PushI Ly, "Ele Txt AATy"
PushI Ly, "Ele Loc Loc"
PushI Ly, "Ele Expr Expr"
PushI Ly, "Ele Mem Rmk"
PushI Ly, "Fld Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
PushI Ly, "Fld Expr Txt [Expr=Loc & 'abc']"
PushI Ly, "Des A.     AA BB "
PushI Ly, "Des A.     CC DD "
PushI Ly, "Des .ANm   AA BB "
PushI Ly, "Des A.ANm  TF_Des-AA-BB"
Set EptB.EF.E = New Dictionary
Set EptB.EF.F = New Dictionary
Set EptB.Des = New Dictionary
GoTo Tst
Tst:
    ActB = New_StruBase(Ly)
    If Not StruBaseIsEq(ActB, EptB) Then Stop
    Return
End Sub

Function StruBaseIsEq(A As StruBase, B As StruBase) As Boolean
If Not Dic_IsEq(A.EF.E, B.EF.E) Then Exit Function
If Not Dic_IsEq(A.EF.F, B.EF.F) Then Exit Function
If Not Dic_IsEq(A.Des, B.Des) Then Exit Function
End Function

