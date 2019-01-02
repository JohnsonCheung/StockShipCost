Attribute VB_Name = "MDao__Tmp"
Option Compare Binary
Option Explicit
Property Get TmpTd() As DAO.TableDef
Dim O() As DAO.Field2
PushObj O, New_Fd("F1")
'Set TmpTd = NewTd("Tmp", O)
End Property

Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
Dim Fb$
Fb = TmpFb(Fdr, Fnn)
Fb_XCrt Fb
Set TmpDb = Fb_Db(Fb)
End Function
