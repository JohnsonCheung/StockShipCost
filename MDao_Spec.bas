Attribute VB_Name = "MDao_Spec"
Option Compare Binary
Option Explicit

Sub Db_XImp_Spec(A As Database, Spnm)
Dim Ft$
    Ft = Spnm_Ft(Spnm)
    
Dim NoCur As Boolean
Dim NoLas As Boolean
Dim CurOld As Boolean
Dim CurNew As Boolean
Dim SamTim As Boolean
Dim DifSz As Boolean
Dim SamSz As Boolean
Dim DifFt As Boolean
Dim Rs As DAO.Recordset
    
    Q = QQ_Fmt("Select SpecNm,Ft,Lines,Tim,Sz,LdTim from Spec where SpecNm = '?'", Spnm)
    Set Rs = CurDb.OpenRecordset(Q)
    NoCur = Not Ffn_Exist(Ft)
    NoLas = Rs_IsNoRec(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdDTim$
    CurS = Ffn_Sz(Ft)
    CurT = Ffn_Tim(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Sz, -1)
            LasT = Nz(!Tim, 0)
            LasFt = Nz(!Ft, "")
            LdDTim = Dte_DTim(!LdTim)
        End With
    End If
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
    SamSz = CurS = LasS
    DifSz = Not SamSz
    DifFt = Ft <> LasFt
    

Const Imported$ = "***** IMPORTED ******"
Const NoImport$ = "----- no import -----"
Const NoCur____$ = "No Ft."
Const NoLas____$ = "No Last."
Const FtDif____$ = "Ft is dif."
Const SamTimSz__$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld__$ = "Cur is old."
Const CurIsNew__$ = "Cur is new."
Const C$ = "|[SpecNm] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Sz] [Las-Sz] [Imported-Time]."

Dim Dr()
Dr = Array(Spnm, Ft, Ft_Lines(Ft), CurT, CurS, Now)
Select Case True
Case NoCur, SamTim:
Case NoLas: Dr_XIns_Rs Dr, Rs
Case DifFt, CurNew: Dr_XUpd_Rs Dr, Rs
Case Else: Stop
End Select

Dim Av()
Av = Array(Spnm, Db_Nm(A), Ft, LasFt, CurT, LasT, CurS, LasS, LdDTim)
Select Case True
Case NoCur:            XDmp_Lin_AV CSub, NoImport & NoCur____ & C, Av
Case NoLas:            XDmp_Lin_AV CSub, Imported & NoLas____ & C, Av
Case DifFt:            XDmp_Lin_AV CSub, Imported & FtDif____ & C, Av
Case SamTim And SamSz: XDmp_Lin_AV CSub, NoImport & SamTimSz__ & C, Av
Case SamTim And DifSz: XDmp_Lin_AV CSub, NoImport & SamTimDifSz & C, Av
Case CurOld:           XDmp_Lin_AV CSub, NoImport & CurIsOld__ & C, Av
Case CurNew:           XDmp_Lin_AV CSub, Imported & CurIsNew__ & C, Av
Case Else: Stop
End Select
End Sub

Property Get SpecPth$()
SpecPth = Pth_XEns(CurDb_Pth & "Spec\")
End Property

Sub SpecPth_XBrw()
Pth_XBrw SpecPth
End Sub

Sub SpecPth_XClr()
Pth_XClr SpecPth
End Sub

Property Get SpecSchmy() As String()
SpecSchmy = SplitCrLf(SpecSchmLines)
End Property

Sub Db_XEns_SpecTbl(A As Database)
If Not Dbt_Exist(A, "Spec") Then Db_XCrt_SpecTbl A
End Sub

Sub Db_XCrt_SpecTbl(A As Database)

End Sub

Sub XCrt_SpecTbl()
Db_XCrt_SpecTbl CurDb
End Sub

Sub SpecXEnsTbl()
Db_XEns_SpecTbl CurDb
End Sub

Sub Spec_XExp()
SpecPth_XClr
Dim X
For Each X In AyNz(SpecNy)
    Spnm_XExp X
Next
End Sub

Property Get SpecNy() As String()
SpecNy = Db_SpecNy(CurDb)
End Property

Function Db_SpecNy(A As DAO.Database) As String()
Db_SpecNy = Dbtf_StrCol(A, "Spec", "SpecNm")
End Function


