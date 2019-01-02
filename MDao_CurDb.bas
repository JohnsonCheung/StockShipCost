Attribute VB_Name = "MDao_CurDb"
Option Compare Binary
Option Explicit
Const CMod$ = "MDao_CurDb."

Sub CurDb_XCls()
On Error Resume Next
CurAcs.CloseCurrentDatabase
End Sub

Property Get CnStrAy() As String()
CnStrAy = CurDb_CnStrAy
End Property

Property Get CurAcs() As Access.Application
Set CurAcs = Access.Application
End Property

Property Get CurDb() As Database
Const CSub$ = CMod & "CurDb"
Dim O As Database
Set O = CurAcs.CurrentDb
If IsNothing(O) Then XThw_Msg CSub, "No currentDb"
Set CurDb = O
End Property

Property Get CurDb_XChk__Idx() As String()
CurDb_XChk__Idx = Db_XChk__Idx(CurDb)
End Property

Property Get CurDb_CnStrAy() As String()
CurDb_CnStrAy = Db_CnStrAy(CurDb)
End Property

Sub CurDb_XDrp_AllLnkTbl()
Db_XDrp_AllLnkTbl CurDb
End Sub

Property Get CurDb_XLnk_Tny() As String()
CurDb_XLnk_Tny = Db_XLnk_Tny(CurDb)
End Property

Sub CurDbPth_XBrw()
Pth_XBrw CurDb_Pth
End Sub

Property Get CurDb_TdSclLy() As String()
CurDb_TdSclLy = Db_TdSclLy(CurDb)
End Property
Sub CurDb_Pth_XBrw()
Pth_XBrw CurDb_Pth
End Sub
Property Get CurDb_Pth$()
CurDb_Pth = Ffn_Pth(CurDb.Name)
End Property
Property Get CurDb_Stru$()
CurDb_Stru = Db_Stru(CurDb)
End Property

Property Get CurDb_Tny() As String()
CurDb_Tny = Db_Tny(CurDb)
End Property

Property Get CurFb$()
On Error Resume Next
CurFb = CurDb.Name
End Property

Function DftDb(A As Database) As Database
If IsNothing(A) Then
   Set DftDb = CurDb
Else
   Set DftDb = A
End If
End Function

Function Fb_CurDb(A) As Database
CurDb_XCls
CurAcs.OpenCurrentDatabase A
Set Fb_CurDb = CurDb
End Function


Sub XRfh_TmpTbl()
Tbl_XDrp "Tmp"
Db_XAdd_TmpTbl CurDb
End Sub

Property Get Tny() As String()
Tny = CurDb_Tny
End Property

Private Sub Z()
Exit Sub
'CurDb_XCls
'CnSy
'CurAcs
'CurDb
'CurDb_XChk_
'CurDb_CnStrAy
'CurDb_XDrp_AllLnkTbl
'CurDbLnkTny
'CurDbPth
'CurDbPth_XBrw
'CurDb_TdSclLy
'CurDbStru
'CurDb_Tny
'CurFb
'DftDb
'Fb_CurDb
'Tbl_Exist
'OpnCurDb
'RfhTmpTbl
'Tny
End Sub
