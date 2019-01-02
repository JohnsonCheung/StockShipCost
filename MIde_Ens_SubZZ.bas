Attribute VB_Name = "MIde_Ens_SubZZ"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Ens_SubZZ."

Function Arg_ArgSfx$(A)
Dim B$
B = XRmv_Nm(XRmv_PfxSpc(XRmv_PfxSpc(A, "Optional"), "ParamArray"))
Arg_ArgSfx = RTrim(XTak_BefOrAll(B, "=", NoTrim:=True))
End Function

Function Md_SubZZ_EPT$(A As CodeModule) ' SubZZ is Sub ZZ() bodyLines
'Sub ZZ() has all calling of public method with dummy parameter so that it can XShf_-F2
Dim Dcl$()        ' Mth Dcl PUB
Dim PrpGetASet As ASet
Dim Pm$()         ' From Dcl    Dcl & mPm same sz     ' Pm is the string in the bracket of the MthLin
Dim MthNm$()      ' From Dcl    Dcl & mMthNm same sz
Dim Arg$()          ' Each Arg in Arg become on Pm   Eg, 1-Arg = A$, B$, C%, D As XYZ => 4-Pm
                     ' ArgSfxDic is Key=ArgSfx and Val=A, B, C
                     ' ArgSfx is Arg-without-Nm
Dim ArgSfx$()
Dim ArgASet As ASet
Dim ArgDic As Dictionary
Dim CallingPmAy$()
    Dcl = Md_MthLinAy_Pub(A)
    Set PrpGetASet = WPrpGetASet(Dcl)
    Pm = Ay_XTak_BetBkt(Dcl) ' Each Mth return an-Ele to call
    MthNm = MthLinAy_MthNy(Dcl)
    Arg = WArgAy(Pm)
    ArgSfx = WArgSfxAy(Arg)
    Set ArgASet = New_ASet(ArgSfx)
    Set ArgDic = ASet_AbcDic(ArgASet)
    CallingPmAy = WCallingPmAy(Pm, ArgDic)
    
'-------------
Dim O_DimLy$()
Dim O_CallingLy$()
    O_DimLy = WDimLy(ArgDic)   ' 1-Arg => 1-DimLin
    O_CallingLy = WCallingLy(MthNm, Pm, ArgDic, PrpGetASet)

Dim O$()
    PushI O, "Private Sub ZZ()"
    PushIAy O, O_DimLy
    PushIAy O, O_CallingLy
    PushI O, "End Sub"
Md_SubZZ_EPT = JnCrLf(O)
End Function

Function MthLinAy_MthNy(A$()) As String()
Const CSub$ = CMod & "MthLinAy_MthNy"
Dim I, MthNm$, J%
For Each I In AyNz(A)
    MthNm = Lin_MthNm(I)
    If MthNm = "" Then XThw CSub, "Given MthLinAy does not have MthNm", "[MthLin with error] Ix MthLinAy", I, J, A
    PushI MthLinAy_MthNy, MthNm
    J = J + 1
Next
End Function

Private Function WArgAy(PmAy$()) As String()
Dim Pm, Arg
Dim O As New ASet 'Pm the the string in the bracket of the MthLin
                  'Arg is the one Arg in Pm.  Eg 1-Pm: A$, B$ => 2-Arg
                  'ArgDic the Arg mapping to A B C.  Fst Arg will be A Snd Arg will be B, ..
For Each Pm In AyNz(PmAy)
    For Each Arg In AyNz(AyTrim(SplitComma(Pm)))
        PushI WArgAy, Arg
    Next
Next
End Function

Private Function WArgSfxAy(ArgAy$()) As String()
Dim Arg
For Each Arg In AyNz(ArgAy)
    PushI WArgSfxAy, Arg_ArgSfx(Arg)
Next
End Function

Private Function WCallingLin$(MthNm, CallingPm$, PrpGetASet As ASet)
If ASet_XHas(PrpGetASet, MthNm) Then
    WCallingLin = "XX = " & MthNm & "(" & CallingPm & ")"  ' The MthNm is object, no need to add [Set] XX =, the compiler will not check for this
Else
    WCallingLin = MthNm & XAdd_PfxSpc_IfAny(CallingPm)
End If
End Function

Private Function WCallingLy(MthNy$(), PmAy$(), ArgDic As Dictionary, PrpGetASet As ASet) As String()
'A$() & PmAy$() are same sz
'ArgDic: Key is ArgSfx(Arg-without-Name), Val is A,B,..
'CallingLin is {MthNm} A,B,C,...
'PrpGetASet    is PrpNm set
Dim MthNm, CallingPm$, Pm$, J%, O$()
For Each MthNm In AyNz(MthNy)
    Pm = PmAy(J)
    CallingPm = WCallingPm(Pm, ArgDic)
    PushI O, WCallingLin(MthNm, CallingPm, PrpGetASet)
    J = J + 1
Next
WCallingLy = AyQSrt(O)
End Function

Private Function WCallingPm$(Pm, ArgDic As Dictionary)
Dim O$(), Arg
For Each Arg In AyNz(AyTrim(SplitComma(Pm)))
    PushI O, ArgDic(Arg_ArgSfx(Arg))
Next
WCallingPm = JnCommaSpc(O)
End Function

Private Function WCallingPmAy(PmAy$(), ArgDic As Dictionary) As String()
Dim Pm
For Each Pm In AyNz(PmAy)
    PushI WCallingPmAy, WCallingPm(Pm, ArgDic)
Next
End Function

Private Function WDimLy(ArgDic As Dictionary) As String() '1-Arg => 1-DimLin
Dim ArgSfx
For Each ArgSfx In ArgDic.Keys
    PushI WDimLy, "Dim " & ArgDic(ArgSfx) & ArgSfx
Next
PushI WDimLy, "Dim XX" 'For Prp
End Function

Private Function WPrpGetASet(MthLinAy$()) As ASet
Dim Lin, O As ASet
Set O = EmpASet
For Each Lin In AyNz(MthLinAy)
    If Lin_IsPrp(Lin) Then ASet_XPush O, Lin_MthNm(Lin)
Next
Set WPrpGetASet = O
End Function

Private Sub Z_Md_SubZZ_EPT()
Dim M As CodeModule
GoSub T1
'GoSub T2
Exit Sub
T1:
    Set M = Md("MDao_Z_Db_Dbt")
    GoTo Tst
T2:
    Set M = CurMd
    GoTo Tst
Tst:
    Act = Md_SubZZ_EPT(M)
    Brw Act
    Stop
    C
Return
End Sub

Private Sub ZZ()
Dim A
Dim B As CodeModule
Dim C$()
Dim XX
Arg_ArgSfx A
Md_SubZZ_EPT B
MthLinAy_MthNy C
End Sub

Private Sub Z()
Z_Md_SubZZ_EPT
End Sub
