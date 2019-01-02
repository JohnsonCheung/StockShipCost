Attribute VB_Name = "MIde_Dcl_EnmAndTy"
Option Compare Binary
Option Explicit
Const CMod$ = "MIde_Dcl_EnmAndTy."

Function Dcl_EnmBdyLy(A$(), EnmNm$) As String()
Const CSub$ = CMod & "Dcl_EnmBdyLy"
Dim B%: B = Dcl_EnmIx(A, EnmNm): If B = -1 Then XThw CSub, "No EnmNm in Src", "EnmNm Src", EnmNm, A
Dim J%
For J = B To UB(A)
   PushI Dcl_EnmBdyLy, A(J)
   If XHas_Pfx(A(J), "End Enum") Then Exit Function
Next
XThw CSub, "No End Enum", "Src EnmNm", A, EnmNm
End Function

Function Dcl_EnmIx%(A$(), EnmNm$)
Dim J%, L
For Each L In AyNz(A)
    L = XRmv_Mdy(L)
    If XShf_XEnm(L) Then
        If XTak_Nm(L) = EnmNm Then
            Dcl_EnmIx = J
            Exit Function
        End If
    End If
    If Lin_IsMth(L) Then Exit For
    J = J + 1
Next
Dcl_EnmIx = -1
End Function

Function Dcl_EnmNy(A$()) As String()
Dim L
For Each L In AyNz(A)
   PushNonBlankStr Dcl_EnmNy, LinEnmNm(L)
Next
End Function

Function Any_Dcl_TyNm(A$(), TyNm$) As Boolean
Dim L
For Each L In AyNz(A)
    If Lin_TyNm(L) = TyNm Then Any_Dcl_TyNm = True: Exit Function
Next
End Function

Function Dcl_NEnum%(A$())
Dim L, O%
For Each L In AyNz(A)
   If IsEmnLin(L) Then O = O + 1
Next
Dcl_NEnum = O
End Function

Function DclTyNm_TyFmIx%(A$(), TyNm$)
Dim J%, L$
For J = 0 To UB(A)
   If Lin_TyNm(A(J)) = TyNm Then DclTyNm_TyFmIx = J: Exit Function
Next
DclTyNm_TyFmIx = -1
End Function

Function DclTyNm_TyFTIx(A$(), TyNm$) As FTIx
Dim FmI&: FmI = DclTyNm_TyFmIx(A, TyNm)
Dim ToI&: ToI = DclTyToIx(A, FmI)
Set DclTyNm_TyFTIx = New_FTIx(FmI, ToI)
End Function

Function DclTyIx_TyToIx%(A$(), TyIx%)
If 0 > TyIx Then DclTyIx_TyToIx = -1: Exit Function
Dim O&
For O = TyIx + 1 To UB(A)
   If XHas_Pfx(A(O), "End Type") Then DclTyIx_TyToIx = O: Exit Function
Next
DclTyIx_TyToIx = -1
End Function

Function DclTyNm_TyLines$(A$(), TyNm$)
DclTyNm_TyLines = JnCrLf(DclTyNm_TyLy(A, TyNm))
End Function

Function DclTyNm_TyLy(A$(), TyNm$) As String()
DclTyNm_TyLy = Ay_XWh_FTIx(A, DclTyNm_TyFTIx(A, TyNm))
End Function

Function DclTyNmIx&(A$(), TyNm)
Dim J%
For J = 0 To UB(A)
   If Lin_TyNm(A(J)) = TyNm Then DclTyNmIx = J: Exit Function
Next
DclTyNmIx = -1
End Function

Function DclTyNy(A$()) As String()
Dim L
For Each L In AyNz(A)
    PushNonBlankStr DclTyNy, Lin_TyNm(L)
Next
End Function

Private Function DclTyToIx%(A$(), FmIx)
If 0 > FmIx Then DclTyToIx = -1: Exit Function
Dim O&
For O = FmIx + 1 To UB(A)
   If XHas_Pfx(A(O), "End Type") Then DclTyToIx = O: Exit Function
Next
DclTyToIx = -1
End Function

Function IsEmnLin(A) As Boolean
IsEmnLin = XHas_Pfx(XRmv_Mdy(A), "Enum ")
End Function

Function IsTyLin(A) As Boolean
IsTyLin = XHas_Pfx(XRmv_Mdy(A), "Type ")
End Function

Function LinEnmNm$(A)
Dim L$: L = XRmv_Mdy(A)
If XShf_XEnm(L) Then LinEnmNm = XTak_Nm(L)
End Function

Function Lin_TyNm$(A)
Dim L$: L = XRmv_Mdy(A)
If XShf_T(L) Then Lin_TyNm = XTak_Nm(L)
End Function

Function MdEnmBdyLy(A As CodeModule, EnmNm$) As String()
MdEnmBdyLy = Dcl_EnmBdyLy(Md_DclLy(A), EnmNm)
End Function

Function MdEnmMbrCnt%(A As CodeModule, EnmNm$)
MdEnmMbrCnt = Sz(MdEnmMbrLy(A, EnmNm))
End Function

Function MdEnmMbrLy(A As CodeModule, EnmNm$) As String()
MdEnmMbrLy = Ay_XWh_CdLin(MdEnmBdyLy(A, EnmNm))
End Function

Function MdEnmNy(A As CodeModule) As String()
MdEnmNy = Dcl_EnmNy(Md_DclLy(A))
End Function

Function MdNEnm%(A As CodeModule)
MdNEnm = Dcl_NEnum(Md_DclLy(A))
End Function

Function Md_TyLCC(A As CodeModule, TyNm$) As LCC
Dim R&, C1&, C2&
R = Md_TyLno(A, TyNm)
If R > 0 Then
    With SubStrPos(A.Lines(R, 1), TyNm)
        C1 = .FmIx
        C2 = .ToIx
    End With
End If
Md_TyLCC = LCC(R, C1, C2)
End Function

Function Md_TyLno$(A As CodeModule, TyNm$)
Md_TyLno = -1
End Function

Function Md_TyNm$(A As CodeModule)
Md_TyNm = CmpTy_ShtNm(Md_CmpTy(A))
End Function

Function Md_TyNy(A As CodeModule) As String()
Md_TyNy = Ay_XSrt(DclTyNy(Md_DclLy(A)))
End Function

Function PjTyNy(A As VBProject, Optional TyNmPatn$ = ".", Optional MdNmPatn$ = ".") As String()
Dim I, Ny$(), O$()
For Each I In AyNz(Pj_MdAy(A, WhMd(Nm:=WhNm(MdNmPatn))))
    Ny = Md_TyNy(CvMd(I))
    Ny = Ay_XWh_Patn(Ny, TyNmPatn)
    PushIAy O, Ay_XAdd_Pfx(Ny, Md_Nm(CvMd(I)) & ".")
Next
PjTyNy = AyQSrt(O)
End Function

Function XShf_XEnm(O) As Boolean
XShf_XEnm = XShf_X(O, "Enum")
End Function

Function XShf_XTy(O) As Boolean
XShf_XTy = XShf_X(O, "Type")
End Function


Private Sub Z_DclTyNm_TyLines()
Debug.Print DclTyNm_TyLines(Md_DclLy(CurMd), "AA")
End Sub

Private Sub Z()
Z_DclTyNm_TyLines
MIde_Dcl_EnmAndTy:
End Sub
