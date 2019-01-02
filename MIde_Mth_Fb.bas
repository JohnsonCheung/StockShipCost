Attribute VB_Name = "MIde_Mth_Fb"
Option Compare Binary
Option Explicit
Public Const MthLocFx$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.xlsx"
Public Const MthFb$ = "C:\Users\User\Desktop\Vba-Lib-1\Mth.accdb"
Public Const WrkFb$ = "C:\Users\User\Desktop\Vba-Lib-1\MthWrk.accdb"
Sub XEns_MthTbl()
Dim A As Drs
Set A = CurVbe_MthDrs
Drs_XRpl_Dbt A, MthDb, "Mth"
End Sub

Sub XEns_MthFb()
MthFb_XEns MthFb
End Sub

Function MthFb_XEns(A$) As Database
Fb_XEns A
Dim Db As Database
Set Db = Fb_Db(A)
DbSchm_XEns Db, MthSchm
End Function

Property Get MthDb() As Database
Static A As Database, B As Boolean
If Not B Then
    B = True
    Set A = Fb_Db(MthFb)
End If
Set MthDb = A
End Property

Sub XBrw_MthFb()
Fb_XBrw MthFb
End Sub

Private Property Get MthSchm$()
Const A_1$ = "Ele Nm  Md Pj" & _
vbCrLf & "Ele T50 MchStr" & _
vbCrLf & "ELe T10 MthPfx" & _
vbCrLf & "Ele Txt Pj_Ffn Prm Ret LinRmk" & _
vbCrLf & "Ele T3  Ty Mdy" & _
vbCrLf & "Ele T4  Md_Ty" & _
vbCrLf & "Ele Lng Lno" & _
vbCrLf & "Ele Mem Lines TopRmk" & _
vbCrLf & "Tbl MthCache | Pj_Ffn Md Nm Ty | Mdy Prm Ret LinRmk TopRmk Lines Lno Pj PjDte Md_Ty" & _
vbCrLf & "Tbl MthLoc   | Nm             | MthMchTy MchStr ToMdNm" & _
vbCrLf & "Tbl MthPfxMd | MthPfx         | MdNm" & _
vbCrLf & ""

MthSchm = A_1
End Property

