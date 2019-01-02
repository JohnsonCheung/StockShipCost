Attribute VB_Name = "MIde_Z_Pj_Rf_Add"
Option Explicit
Option Compare Binary
Const CMod$ = "MIde_Z_Pj_Rf_Add."

Sub Pj_XSet_Rf_ByRfFfn(A As VBProject, RfNm, RfFfn)
Const CSub$ = CMod & "Pj_XSet_Rf_ByRfFfn"
If Pj_XHas_Rf(A, RfNm) Then Exit Sub
A.References.AddFromFile RfFfn
FunMsgNyAp_XDmp CSub, "Rf is added to Pj", "Rf Pj RfFfn", RfNm, Pj_Nm(A), RfFfn
End Sub

Sub Pj_XSet_Rf_ByRfFfnFfnAy(A As VBProject, RfFfnAy$())
Dim F
For Each F In RfFfnAy
    If Not Pj_XHas_RfFfn(A, CStr(F)) Then
        A.References.AddFromFile F
    End If
Next
End Sub

Sub Pj_XSet_Rf_ByRfFfn_GUID(A As VBProject, RfNm, RfGUID$, Mjr&, Mnr&)
Const CSub$ = CMod & "Pj_XSet_Rf_ByRfFfn_GUID"
If Pj_XHas_RfGUID(A, RfGUID) Then Exit Sub
A.References.AddFromGuid RfGUID, Mjr, Mnr
FunMsgNyAp_XDmp CSub, "Rf is added to Pj", "Rf Pj RfGUID Major Minor", RfNm, Pj_Nm(A), RfGUID, Mjr, Mnr
End Sub

Sub Pj_XSet_Rf_ByRfFfn_GUID_LIN(A As VBProject, GUID_Lin)
Dim GUID$, Maj&, Min&, RfNm$
Lin_3TRstAsg GUID_Lin, RfNm, GUID, Maj, Min
Pj_XSet_Rf_ByRfFfn_GUID A, RfNm, GUID, Maj, Min
End Sub

Sub Pj_XSet_Rf_ByRfFfn_BY_LIST(A As VBProject, RfNy$(), RfFfnAy$())
Dim J%
For J = 0 To UB(RfNy)
    Pj_XSet_Rf_ByRfFfn A, RfNy(J), RfFfnAy(J)
Next
End Sub

Private Sub Pj_XSet_Rf_ByRfFfn_BY_S1S2Ay(A As VBProject, B() As S1S2)
Pj_XSet_Rf_ByRfFfn_BY_LIST A, S1S2Ay_Sy1(B), S1S2Ay_Sy2(B)
End Sub

Function PjRfNyDIC(A As VBProject, RfDic As Dictionary, RfDicNm$) As String()
Const CSub$ = CMod & "PjRfNyDIC"
If Not RfDic.Exists(A.Name) Then
    XThw CSub, "Pj is not defined in RfDic", "Pj RfDicNm RfDic", Pj_Nm(A), RfDicNm, Dic_Fmt(RfDic)
End If
PjRfNyDIC = CvNy(RfDic(A.Name))
End Function

Private Sub ZZ()
Dim A As VBProject
Dim B As Variant
Dim C$()
Dim D$
Dim E&
Dim G As Dictionary
Pj_XSet_Rf_ByRfFfn A, B, B
Pj_XSet_Rf_ByRfFfnFfnAy A, C
Pj_XSet_Rf_ByRfFfn_GUID A, B, D, E, E
Pj_XSet_Rf_ByRfFfn_BY_LIST A, C, C
PjRfNyDIC A, G, D
End Sub

Private Sub Z()
End Sub
