Attribute VB_Name = "MAdoX_CnStr"
Option Compare Binary
Option Explicit
Private Sub Z_Fb_CnStr_Ado()
Dim CnStr$
'
CnStr = Fb_CnStr_Ado(Samp_Fb_Duty_Dta)
GoSub Tst
'
CnStr = Fb_CnStr_Ado(CurrentDb.Name)
'GoSub Tst
Exit Sub
Tst:
    CnStr_Cn(CnStr).Close
    Return
End Sub
Private Sub Z_CnStr_Cn()
Dim O As ADODB.Connection
Set O = CnStr_Cn(Samp_CnStr_ADO_SAMP_SQL_EXPR_NOT_WRK)
Stop
End Sub
Function CnStr_Cn(AdoCnStr) As ADODB.Connection
Set CnStr_Cn = New ADODB.Connection
CnStr_Cn.Open AdoCnStr
End Function

Function Fb_CnStr_Ado$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;" & _
            "Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;" & _
            "Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;" & _
            "Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;" & _
            "Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
'Fb_CnStr_Ado = QQ_Fmt("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
'Locking Mode=1 means page (or record level) according to https://www.spreadsheet1.com/how-to-refresh-pivottables-without-locking-the-source-workbook.html
'The ADO connection object initialization property which controls how the database is locked, while records are being read or modified is: Jet OLEDB:Database Locking Mode
'Please note:
'The first user to open the database determines the locking mode to be used while the database remains open.
'A database can only be opened is a single mode at a time.
'For Page-level locking, set property to 0
'For Row-level locking, set property to 1
'With 'Jet OLEDB:Database Locking Mode = 0', the source spreadshseet is locked, while PivotTables update. If the property is set to 1, the source file is not locked. Only individual records (Table rows) are locked sequentially, while data is being read.
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
'Const C$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=?" 'C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb (It Works)
Fb_CnStr_Ado = QQ_Fmt(C, A)
End Function

Function Fx_CnStr_Ado$(A)
'Fx_CnStr_Ado = QQ_Fmt("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0;HDR=YES""", A)
Fx_CnStr_Ado = QQ_Fmt("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A)
End Function

Function CnStr_DtaSrcVal$(A)
CnStr_DtaSrcVal = TakBet(A, "Data Source=", ";")
End Function
