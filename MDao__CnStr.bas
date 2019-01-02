Attribute VB_Name = "MDao__CnStr"
Option Compare Text
Option Explicit
Function Fb_CnStr_DAO$(A)
Fb_CnStr_DAO = ";DATABASE=" & A & ";"
End Function

Function Fx_CnStr_DAO$(A)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
'Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
Dim O$
Select Case Ffn_Ext(A)
Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & A & ";"
Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & A & ";"
Case Else: Stop
End Select
Fx_CnStr_DAO = O
End Function

Function Fb_CnStr_OleAdo$(A) 'Return a connection used as WbConnection
Fb_CnStr_OleAdo = "OLEDb;" & Fb_CnStr_Ado(A)
End Function

Function Dbt_CnStr$(A As Database, T)
On Error Resume Next
Dbt_CnStr = A.TableDefs(T).Connect
End Function

Function Tbl_CnStr$(T)
Tbl_CnStr = Dbt_CnStr(CurDb, T)
End Function

Function Fb_CnStr_WbCn$(Fb$)
Fb_CnStr_WbCn = Fb_CnStr_OleAdo(Fb)
'Fb_CnStr_WbCn = QQ_Fmt("Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
'Fb_CnStr_WbCn = QQ_Fmt("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
End Function

Function Any_CnStr_Td(A As DAO.TableDef) As Boolean
Any_CnStr_Td = A.Connect <> ""
End Function
